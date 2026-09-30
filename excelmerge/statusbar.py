"""공유 상태바 — 텍스트 메시지 + 우측 진행률 바.

기존에는 긴 작업(로딩·비교·저장·폴더 스캔/병합)이 `showMessage` 텍스트로만 표시됐다.
`QStatusBar` 를 이 클래스로 교체하면 탭들은 기존 `self.status.showMessage(...)` 를 그대로
쓰면서, 추가로 다음 API 로 진행률 바를 제어할 수 있다:

    status.set_progress(done, total)   # total>0 확정, total<=0 바쁨(불확정)
    status.begin_busy()                # 불확정 진행 시작
    status.end_progress()              # 바 숨김

진행률 바는 `addPermanentWidget` 로 우측에 얹은 숨김 위젯이며, 필요할 때만 표시된다.
(업데이터의 QProgressDialog `_prog` 패턴과 동일한 확정/불확정 전환 규칙.)
"""
from PyQt5.QtWidgets import QStatusBar, QProgressBar


class StatusBar(QStatusBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._bar = QProgressBar(self)
        self._bar.setFixedWidth(180)
        self._bar.setMaximumHeight(14)
        self._bar.setTextVisible(False)
        self._bar.hide()
        self.addPermanentWidget(self._bar)

    def set_progress(self, done: int, total: int):
        """확정(total>0) 또는 바쁨(total<=0) 진행률 표시. 자동으로 바를 보인다."""
        if total and total > 0:
            self._bar.setMaximum(int(total))
            self._bar.setValue(max(0, min(int(done), int(total))))
        else:
            self._bar.setMaximum(0)   # 불확정("바쁨") 애니메이션
        if self._bar.isHidden():
            self._bar.show()

    def begin_busy(self):
        """전체량을 모르는 작업 — 불확정 진행 바를 켠다."""
        self.set_progress(0, 0)

    def end_progress(self):
        """진행 바를 숨긴다(작업 완료/에러 시)."""
        self._bar.reset()
        self._bar.hide()
