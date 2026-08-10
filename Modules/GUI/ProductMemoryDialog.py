from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from Modules.Common.ErrorReport import FailureDetails
from Modules.Common.paths import default_download_dir
from Modules.GUI.Dialogs import ErrorReportDialog
from Modules.WMS.ProductMemory import ProductMemory, ProductMemoryRecord


def _format_decimal(value: Decimal | None, *, digits: int | None = None) -> str:
    if value is None:
        return "-"
    if digits is not None:
        text = f"{value:.{digits}f}"
    else:
        text = format(value, "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


class ProductMemoryDialog(QDialog):
    """Browse and transfer the SKU weight/classification memory."""

    memory_changed = Signal()

    FILTERS = ("전체", "경량", "중량", "고단", "양곡", "미분류")

    def __init__(self, memory: ProductMemory, parent=None):
        super().__init__(parent)
        self.memory = memory
        self.setWindowTitle("저장된 상품 분류")
        self.setObjectName("ProductMemoryDialog")
        self.resize(1040, 560)
        self.setMinimumSize(820, 440)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(14)

        title_row = QHBoxLayout()
        title = QLabel("저장된 상품 무게와 분류")
        title.setObjectName("DialogTitle")
        title_row.addWidget(title)
        title_row.addStretch(1)
        filter_label = QLabel("분류")
        filter_label.setObjectName("FieldLabel")
        title_row.addWidget(filter_label)
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(self.FILTERS)
        self.filter_combo.currentTextChanged.connect(self.refresh)
        title_row.addWidget(self.filter_combo)
        layout.addLayout(title_row)

        help_label = QLabel(
            "SKU별 상품 무게는 다음 실행에서 WMS 재조회를 생략합니다. "
            "수동 분류는 메인 표의 분류 버튼으로 변경할 수 있습니다."
        )
        help_label.setWordWrap(True)
        help_label.setObjectName("HelpText")
        layout.addWidget(help_label)

        self.table = QTableWidget(0, 7)
        self.table.setObjectName("StoredProductTable")
        self.table.setHorizontalHeaderLabels(
            ["SKU ID", "상품명", "무게(g)", "분류", "적용 방식", "1팔렛트 중량(kg)", "측정 시각"]
        )
        self.table.setWordWrap(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(40)
        header = self.table.horizontalHeader()
        header.setMinimumHeight(42)
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table, 1)

        buttons = QHBoxLayout()
        self.import_button = QPushButton("가져오기")
        self.export_button = QPushButton("내보내기")
        self.delete_button = QPushButton("선택 삭제")
        self.delete_button.setObjectName("DangerButton")
        close_button = QPushButton("닫기")
        self.import_button.clicked.connect(self._import_records)
        self.export_button.clicked.connect(self._export_records)
        self.delete_button.clicked.connect(self._delete_selected)
        close_button.clicked.connect(self.accept)
        buttons.addWidget(self.import_button)
        buttons.addWidget(self.export_button)
        buttons.addWidget(self.delete_button)
        buttons.addStretch(1)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)

        self.refresh()

    def refresh(self, *_args) -> None:
        selected_filter = self.filter_combo.currentText()
        records = [record for record in self.memory.entries() if self._matches(record, selected_filter)]
        self.table.setRowCount(len(records))
        for row_index, record in enumerate(records):
            category = record.effective_category or "미분류"
            if record.category_override:
                source = "수동"
            elif record.automatic_category:
                source = "자동"
            elif record.weight_grams is not None:
                source = "무게만"
            else:
                source = "측정 전"
            values = (
                record.sku_id,
                record.product_name,
                _format_decimal(record.weight_grams),
                category,
                source,
                _format_decimal(record.pallet_weight_kg, digits=3),
                record.measured_at or "-",
            )
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if column_index == 0:
                    item.setData(Qt.ItemDataRole.UserRole, record.sku_id)
                if column_index == 1:
                    item.setToolTip(record.product_name)
                self.table.setItem(row_index, column_index, item)

    @staticmethod
    def _matches(record: ProductMemoryRecord, selected_filter: str) -> bool:
        if selected_filter == "전체":
            return True
        category = record.effective_category or "미분류"
        return category == selected_filter

    @staticmethod
    def _records_equivalent(
        existing: ProductMemoryRecord,
        incoming: ProductMemoryRecord,
    ) -> bool:
        return (
            existing.sku_id == incoming.sku_id
            and existing.product_name == incoming.product_name
            and existing.weight_grams == incoming.weight_grams
            and existing.automatic_category == incoming.automatic_category
            and existing.category_override == incoming.category_override
            and existing.boxes_per_pallet == incoming.boxes_per_pallet
            and existing.pallet_weight_kg == incoming.pallet_weight_kg
        )

    @staticmethod
    def _record_description(record: ProductMemoryRecord) -> str:
        weight = f"{record.weight_grams}g" if record.weight_grams is not None else "미측정"
        return (
            f"상품명: {record.product_name or '-'}\n"
            f"무게: {weight}\n"
            f"분류: {record.effective_category or '미분류'}"
        )

    def _ask_duplicate_action(
        self,
        existing: ProductMemoryRecord,
        incoming: ProductMemoryRecord,
        *,
        index: int,
        total: int,
    ) -> str:
        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Icon.Question)
        dialog.setWindowTitle("중복 SKU 상품 분류 확인")
        dialog.setText(
            f"SKU {incoming.sku_id}가 이미 저장되어 있습니다. ({index}/{total})\n"
            "가져온 무게와 수동 분류로 덮어쓸까요?"
        )
        dialog.setInformativeText(
            "[현재 저장값]\n"
            f"{self._record_description(existing)}\n\n"
            "[가져올 값]\n"
            f"{self._record_description(incoming)}"
        )
        overwrite_button = dialog.addButton("덮어쓰기", QMessageBox.ButtonRole.AcceptRole)
        keep_button = dialog.addButton("기존 유지", QMessageBox.ButtonRole.RejectRole)
        cancel_button = dialog.addButton(
            "가져오기 취소",
            QMessageBox.ButtonRole.DestructiveRole,
        )
        dialog.setDefaultButton(keep_button)
        dialog.setEscapeButton(cancel_button)
        dialog.exec()
        clicked = dialog.clickedButton()
        if clicked is overwrite_button:
            return "overwrite"
        if clicked is keep_button:
            return "keep"
        return "cancel"

    def _export_records(self) -> None:
        default_name = f"UnHelper_상품분류_{datetime.now():%Y%m%d_%H%M%S}.json"
        selected, _filter = QFileDialog.getSaveFileName(
            self,
            "상품 분류 내보내기",
            str(default_download_dir() / default_name),
            "JSON 파일 (*.json)",
        )
        if not selected:
            return
        destination = Path(selected)
        if destination.suffix.lower() != ".json":
            destination = destination.with_suffix(".json")
        try:
            self.memory.export_to(destination)
        except Exception as exc:
            self._show_error("상품 분류 내보내기 실패", exc)
            return
        QMessageBox.information(self, "내보내기 완료", f"상품 분류 목록을 저장했습니다.\n{destination}")

    def _import_records(self) -> None:
        selected, _filter = QFileDialog.getOpenFileName(
            self,
            "상품 분류 가져오기",
            str(default_download_dir()),
            "JSON 파일 (*.json)",
        )
        if not selected:
            return
        try:
            records = ProductMemory.records_from(selected)
            duplicates = tuple(
                (existing, record)
                for record in records
                if (existing := self.memory.get(record.sku_id)) is not None
                and not self._records_equivalent(existing, record)
            )
            overwrite_sku_ids: set[str] = set()
            for index, (existing, incoming) in enumerate(duplicates, start=1):
                action = self._ask_duplicate_action(
                    existing,
                    incoming,
                    index=index,
                    total=len(duplicates),
                )
                if action == "cancel":
                    return
                if action == "overwrite":
                    overwrite_sku_ids.add(incoming.sku_id)
            summary = self.memory.import_records(
                records,
                overwrite_sku_ids=overwrite_sku_ids,
            )
        except Exception as exc:
            self._show_error("상품 분류 가져오기 실패", exc)
            return
        self.refresh()
        if summary.added or summary.overwritten:
            self.memory_changed.emit()
        QMessageBox.information(
            self,
            "가져오기 완료",
            "상품 분류 가져오기를 완료했습니다.\n"
            f"추가: {summary.added}개\n"
            f"덮어쓰기: {summary.overwritten}개\n"
            f"기존 항목 유지: {summary.skipped}개",
        )

    def _delete_selected(self) -> None:
        sku_ids = {
            str(self.table.item(index.row(), 0).data(Qt.ItemDataRole.UserRole))
            for index in self.table.selectionModel().selectedRows()
            if self.table.item(index.row(), 0) is not None
        }
        if not sku_ids:
            QMessageBox.information(self, "선택 필요", "삭제할 상품을 선택해 주세요.")
            return
        answer = QMessageBox.question(
            self,
            "저장 상품 삭제",
            f"선택한 {len(sku_ids)}개 SKU를 삭제하시겠습니까?\n다음 실행에서는 WMS에서 다시 조회합니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        deleted_count = 0
        try:
            for sku_id in sku_ids:
                if self.memory.delete(sku_id):
                    deleted_count += 1
        except Exception as exc:
            if deleted_count:
                self.refresh()
                self.memory_changed.emit()
            self._show_error("저장 상품 삭제 실패", exc)
            return
        self.refresh()
        if deleted_count:
            self.memory_changed.emit()

    def _show_error(self, title: str, exc: Exception) -> None:
        ErrorReportDialog(
            title,
            FailureDetails.from_exception(exc),
            context={"category": "상품 분류 메모리"},
            parent=self,
        ).exec()
