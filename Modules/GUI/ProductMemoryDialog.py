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

    FILTERS = ("전체", "경량", "중량", "고단", "미분류")

    def __init__(self, memory: ProductMemory, parent=None):
        super().__init__(parent)
        self.memory = memory
        self.setWindowTitle("저장된 상품 분류")
        self.resize(1040, 560)
        self.setMinimumSize(820, 440)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        title_row = QHBoxLayout()
        title = QLabel("저장된 상품 무게와 분류")
        title.setStyleSheet("font-size: 17pt; font-weight: 800;")
        title_row.addWidget(title)
        title_row.addStretch(1)
        title_row.addWidget(QLabel("분류"))
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
        help_label.setStyleSheet("color: #A1A1AA;")
        layout.addWidget(help_label)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(
            ["SKU ID", "상품명", "무게(g)", "분류", "적용 방식", "1팔렛트 중량(kg)", "측정 시각"]
        )
        self.table.setWordWrap(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table, 1)

        buttons = QHBoxLayout()
        self.import_button = QPushButton("가져오기")
        self.export_button = QPushButton("내보내기")
        self.delete_button = QPushButton("선택 삭제")
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
            source = "수동" if record.category_override else ("자동" if record.automatic_category else "측정 전")
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
            summary = self.memory.import_from(selected)
        except Exception as exc:
            self._show_error("상품 분류 가져오기 실패", exc)
            return
        self.refresh()
        if summary.added:
            self.memory_changed.emit()
        QMessageBox.information(
            self,
            "가져오기 완료",
            "상품 분류 가져오기를 완료했습니다.\n"
            f"추가: {summary.added}개\n기존 항목 유지: {summary.skipped}개",
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
