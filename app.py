from pathlib import Path

from config import MyAppConfig, MyAppInputConfig, MyAppOutputConfig
from pyfedappwrap.learning.app_types import BasePrePostProcessApp

from ffq_extractor.batch import clamp_workers, process_input
from ffq_extractor.mapping import default_mapping_path


class FfqExtractorApp(BasePrePostProcessApp[MyAppConfig, MyAppInputConfig, MyAppOutputConfig]):
    def __init__(self):
        super().__init__()

    def run_process(self, data: MyAppInputConfig) -> MyAppOutputConfig:
        archive = Path(data.ffq_archive)
        nutrition_csv = Path(".") / "nutrition.csv"

        workers = clamp_workers(self.config.workers)
        sheet = (self.config.sheet or "").strip()
        date_cell = (self.config.date_cell or "A1").strip() or "A1"
        patient_id_cell = (self.config.patient_id_cell or "A3").strip() or "A3"

        self.logger.info(
            f"FFQ extract start: input={archive} workers={workers} "
            f"sheet={(sheet or '<first>')!r} date_cell={date_cell!r} "
            f"patient_id_cell={patient_id_cell!r}"
        )

        def _progress(done: int, total: int, batch) -> None:
            rate = batch.rows_written / batch.elapsed_seconds if batch.elapsed_seconds else 0
            self.logger.info(
                f"Progress {done}/{total} files_ok={batch.files_ok} "
                f"files_failed={batch.files_failed} rows={batch.rows_written} "
                f"(~{rate:.0f} rows/s)"
            )

        def _on_error(member: str, message: str) -> None:
            self.logger.error(f"FFQ extract failed for {member}: {message}")

        result = process_input(
            archive,
            nutrition_csv,
            mapping_path=default_mapping_path(),
            sheet_name=sheet,
            date_cell=date_cell,
            patient_id_cell=patient_id_cell,
            workers=workers,
            progress_callback=_progress,
            error_callback=_on_error,
        )

        self.logger.info(
            f"FFQ extract done: files_ok={result.files_ok} "
            f"files_failed={result.files_failed} rows={result.rows_written} "
            f"elapsed={result.elapsed_seconds:.2f}s"
        )

        if result.files_failed and result.failed_members:
            preview = ", ".join(result.failed_members[:10])
            self.logger.warning(
                f"{result.files_failed} file(s) failed. Examples: {preview}"
            )

        if not result.ok:
            self.logger.error(
                "Extraction finished with failures "
                f"(ok={result.files_ok}, failed={result.files_failed})"
            )

        if not nutrition_csv.exists():
            nutrition_csv.write_text(
                "Unique Patient ID,Questionnaire date,Portion size,"
                "Frequency consumption,Food group,Food subgroup,Food item\n",
                encoding="utf-8",
            )

        return MyAppOutputConfig(nutrition_csv=nutrition_csv)
