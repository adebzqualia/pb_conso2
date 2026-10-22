import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException
from xlsxwriter.exceptions import XlsxWriterException

from .history import atomic_json, history_lock, load_history, merge_snapshot
from .reader import extract_folder
from .workbook import write_workbook


def _presentation_config(path):
    """Validate the presentation file before creating any output or history."""
    config = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(config, dict) or config.get("schema_version") != 1:
        raise ValueError("Configuration de présentation invalide : schema_version doit valoir 1.")
    if config.get("country_currency", "local") not in ("local", "EUR") or config.get("global_currency", "EUR") != "EUR":
        raise ValueError("Devises de présentation : pays 'local' ou 'EUR', global 'EUR'.")
    for key in ("style", "category_labels", "country_options", "comments"):
        if not isinstance(config.get(key, {}), dict):
            raise ValueError(f"presentation.json : {key} doit être un objet JSON.")
    if not isinstance(config.get("country_order", []), list) or any(not isinstance(eid, str) for eid in config.get("country_order", [])):
        raise ValueError("presentation.json : country_order doit être une liste d'identifiants.")
    if "groups" in config:
        if not isinstance(config["groups"], list):
            raise ValueError("presentation.json : groups doit être une liste.")
        for group in config["groups"]:
            if (not isinstance(group, dict) or not isinstance(group.get("id"), str)
                    or not isinstance(group.get("name"), str)
                    or not isinstance(group.get("entity_ids", []), list)
                    or any(not isinstance(eid, str) for eid in group.get("entity_ids", []))):
                raise ValueError("presentation.json : chaque groupe requiert id, name et entity_ids.")
    for eid, options in config.get("country_options", {}).items():
        if not isinstance(options, dict) or options.get("currency", "local") not in ("local", "EUR"):
            raise ValueError(f"presentation.json : options invalides pour {eid}.")
        if "show_stock_bins" in options and not isinstance(options["show_stock_bins"], bool):
            raise ValueError(f"presentation.json : show_stock_bins doit être booléen pour {eid}.")
    for period, comments in config.get("comments", {}).items():
        if not isinstance(comments, dict):
            raise ValueError(f"presentation.json : comments[{period}] doit être un objet JSON.")
        for eid, text in comments.items():
            if not isinstance(text, str) and (not isinstance(text, list) or any(not isinstance(line, str) for line in text)):
                raise ValueError(f"presentation.json : commentaires invalides pour {period}/{eid}.")
    return config


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Consolidation trimestrielle Positive Balance")
    parser.add_argument("folder", type=Path, help="Dossier QX_20XX contenant les fichiers pays et Taux.xlsx")
    parser.add_argument("--config", type=Path, default=Path(__file__).resolve().parents[1] / "config")
    parser.add_argument("--period", help="Trimestre explicite, par exemple 2026Q2")
    parser.add_argument("--report", choices=("summary", "detailed", "both"), default="summary",
                        help="Version synthétique, détaillée, ou les deux (défaut : summary)")
    parser.add_argument("--output", type=Path,
                        help="Classeur synthétique, ou détaillé si --report detailed")
    parser.add_argument("--detailed-output", type=Path,
                        help="Classeur détaillé (défaut : Synthese_detaillee_2026Q2.xlsx)")
    parser.add_argument("--presentation-config", type=Path,
                        help="Configuration détaillée (défaut : dossier --config/presentation.json)")
    parser.add_argument("--history", type=Path, default=Path("data/history.json"))
    parser.add_argument("--validate-only", action="store_true", help="Lire et contrôler sans écrire")
    parser.add_argument("--strict", action="store_true", help="Arrêter aussi sur les avertissements")
    args = parser.parse_args(argv)
    temporaries = []
    try:
        detailed = args.report in ("detailed", "both")
        if not detailed and (args.detailed_output is not None or args.presentation_config is not None):
            raise ValueError("--detailed-output et --presentation-config nécessitent --report detailed ou both.")
        if args.report == "detailed" and args.output is not None and args.detailed_output is not None:
            raise ValueError("Avec --report detailed, utiliser soit --output soit --detailed-output, pas les deux.")
        presentation_path = (args.presentation_config or args.config / "presentation.json").resolve() if detailed else None
        presentation = _presentation_config(presentation_path) if detailed else None
        snapshot = extract_folder(args.folder.resolve(), args.config.resolve(), args.period)
        issues = snapshot.get("issues", [])
        for issue in issues:
            location = f" [{issue['source']}]" if issue.get("source") else ""
            print(f"{issue['severity'].upper()} {issue['code']} {issue.get('entity', '')}: {issue['message']}{location}")
        if any(i["severity"].lower() == "error" for i in issues):
            raise ValueError("Des données indispensables sont invalides. Corriger les erreurs avant consolidation.")
        if args.strict and issues:
            raise ValueError("Mode strict : des contrôles ou exclusions restent à examiner.")
        if detailed:
            # Validate the model/configuration too, including group membership,
            # without requiring the detailed workbook writer in summary mode.
            from .detail_data import build_presentation_data
            build_presentation_data({"snapshots": {snapshot["period"]: snapshot}}, snapshot["period"], presentation)
        if args.validate_only:
            print(f"Validation terminée : {snapshot['period']}, {len(snapshot['entities'])} entités, {len(issues)} observations.")
            return 0
        outputs = []
        if args.report in ("summary", "both"):
            summary_path = (args.output or Path("outputs") / f"Synthese_{snapshot['period']}.xlsx").resolve()
            outputs.append(("summary", summary_path))
        if detailed:
            if args.report == "detailed":
                detailed_path = args.detailed_output or args.output or Path("outputs") / f"Synthese_detaillee_{snapshot['period']}.xlsx"
            else:
                detailed_path = args.detailed_output or summary_path.parent / f"Synthese_detaillee_{snapshot['period']}.xlsx"
            outputs.append(("detailed", detailed_path.resolve()))
        history_path = args.history.resolve()
        if any(output.suffix.lower() != ".xlsx" for _, output in outputs):
            raise ValueError("La sortie doit porter l'extension .xlsx.")
        # Output/state must never replace input workbooks, configs or one another.
        protected = {p.resolve() for p in args.folder.iterdir() if p.is_file()}
        protected.update(p.resolve() for p in args.config.rglob("*") if p.is_file())
        if presentation_path is not None:
            protected.add(presentation_path)
        output_paths = [output for _, output in outputs]
        if len(set(output_paths)) != len(output_paths):
            raise ValueError("Les deux classeurs de sortie doivent avoir des chemins distincts.")
        if (any(output in protected or output == history_path for output in output_paths)
                or history_path in protected):
            raise ValueError("La sortie/l'historique ne peut pas écraser un fichier source ou de configuration.")
        if detailed:
            snapshot.setdefault("configurations", []).append({"file": str(presentation_path),
                "sha256": hashlib.sha256(presentation_path.read_bytes()).hexdigest()})
        snapshot["generated_utc"] = datetime.now(timezone.utc).isoformat()
        with history_lock(history_path):
            history = load_history(history_path)
            currencies = {e["id"]: e["currency"] for s in history["snapshots"].values() for e in s["entities"]}
            if any(e["id"] in currencies and currencies[e["id"]] != e["currency"] for e in snapshot["entities"]):
                raise ValueError("La devise d'une entité diffère de l'historique. Utiliser un identifiant distinct pour ce changement de périmètre.")
            previous = history["snapshots"].get(snapshot["period"])
            if previous:
                old_ids = {e["id"] for e in previous["entities"]}
                new_ids = {e["id"] for e in snapshot["entities"]}
                if old_ids != new_ids:
                    raise ValueError("Le périmètre de ce trimestre diffère du snapshot enregistré. "
                                     "Utiliser un historique distinct pour un changement de périmètre.")
            history = merge_snapshot(history, snapshot)
            staged = []
            for kind, output in outputs:
                output.parent.mkdir(parents=True, exist_ok=True)
                fd, temporary_name = tempfile.mkstemp(prefix=".positive_balance_", suffix=".xlsx", dir=output.parent)
                os.close(fd)
                temporary = Path(temporary_name)
                temporaries.append(temporary)
                if kind == "summary":
                    write_workbook(temporary, history, snapshot["period"])
                else:
                    from .detailed_workbook import write_detailed_workbook
                    write_detailed_workbook(temporary, history, snapshot["period"], presentation)
                staged.append((temporary, output))
            # Finish generating every requested workbook before replacing any
            # existing output. Each replacement is atomic on its filesystem.
            for temporary, output in staged:
                os.replace(temporary, output)
                temporaries.remove(temporary)
            atomic_json(history_path, history)
        for kind, output in outputs:
            print(f"Classeur {'détaillé' if kind == 'detailed' else 'synthétique'} : {output}")
        print(f"Historique : {history_path}")
        print(f"{len(snapshot['entities'])} entités consolidées. {len(issues)} observations dans Controles.")
        return 0
    except (ValueError, OSError, KeyError, BadZipFile, InvalidFileException, XlsxWriterException) as exc:
        print(f"ERREUR : {exc}", file=sys.stderr)
        return 2
    finally:
        for temporary in temporaries:
            temporary.unlink(missing_ok=True)
