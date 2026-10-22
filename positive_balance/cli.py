import argparse
from datetime import datetime, timezone
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


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Consolidation trimestrielle Positive Balance")
    parser.add_argument("folder", type=Path, help="Dossier QX_20XX contenant les fichiers pays et Taux.xlsx")
    parser.add_argument("--config", type=Path, default=Path(__file__).resolve().parents[1] / "config")
    parser.add_argument("--period", help="Trimestre explicite, par exemple 2026Q2")
    parser.add_argument("--output", type=Path, help="Classeur généré (défaut : outputs/Synthese_2026Q2.xlsx)")
    parser.add_argument("--history", type=Path, default=Path("data/history.json"))
    parser.add_argument("--validate-only", action="store_true", help="Lire et contrôler sans écrire")
    parser.add_argument("--strict", action="store_true", help="Arrêter aussi sur les avertissements")
    args = parser.parse_args(argv)
    temporary = None
    try:
        snapshot = extract_folder(args.folder.resolve(), args.config.resolve(), args.period)
        issues = snapshot.get("issues", [])
        for issue in issues:
            location = f" [{issue['source']}]" if issue.get("source") else ""
            print(f"{issue['severity'].upper()} {issue['code']} {issue.get('entity', '')}: {issue['message']}{location}")
        if any(i["severity"].lower() == "error" for i in issues):
            raise ValueError("Des données indispensables sont invalides. Corriger les erreurs avant consolidation.")
        if args.strict and issues:
            raise ValueError("Mode strict : des contrôles ou exclusions restent à examiner.")
        if args.validate_only:
            print(f"Validation terminée : {snapshot['period']}, {len(snapshot['entities'])} entités, {len(issues)} observations.")
            return 0
        output = (args.output or Path("outputs") / f"Synthese_{snapshot['period']}.xlsx").resolve()
        history_path = args.history.resolve()
        if output.suffix.lower() != ".xlsx":
            raise ValueError("La sortie doit porter l'extension .xlsx.")
        # Output/state must never replace input workbooks, configs or one another.
        protected = {p.resolve() for p in args.folder.iterdir() if p.is_file()}
        protected.update(p.resolve() for p in args.config.rglob("*") if p.is_file())
        if output in protected or history_path in protected or output == history_path:
            raise ValueError("La sortie/l'historique ne peut pas écraser un fichier source ou de configuration.")
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
            output.parent.mkdir(parents=True, exist_ok=True)
            fd, temporary = tempfile.mkstemp(prefix=".positive_balance_", suffix=".xlsx", dir=output.parent)
            os.close(fd)
            write_workbook(Path(temporary), history, snapshot["period"])
            os.replace(temporary, output)
            temporary = None
            atomic_json(history_path, history)
        print(f"Classeur : {output}\nHistorique : {history_path}")
        print(f"{len(snapshot['entities'])} entités consolidées. {len(issues)} observations dans Controles.")
        return 0
    except (ValueError, OSError, KeyError, BadZipFile, InvalidFileException, XlsxWriterException) as exc:
        print(f"ERREUR : {exc}", file=sys.stderr)
        return 2
    finally:
        if temporary:
            Path(temporary).unlink(missing_ok=True)
