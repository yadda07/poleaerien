# -*- coding: utf-8 -*-
from .async_tasks import AsyncTaskBase
from .unified_report import generate_unified_report


class UnifiedReportExportTask(AsyncTaskBase):
    def __init__(self, request):
        super().__init__("Export rapport batch", request)

    def execute(self):
        if self.isCanceled():
            return False

        filepath = generate_unified_report(
            self.params.get('batch_results') or {},
            self.params.get('export_dir', ''),
            {
                'include_comac_drawings': bool(self.params.get('include_comac_drawings')),
                'include_data_dictionary': bool(self.params.get('include_data_dictionary')),
                'progress': self.emit_progress,
                'message': self.emit_message,
                'is_cancelled': self.isCanceled,
                'sro': self.params.get('sro', ''),
            }
        )

        if self.isCanceled() or not filepath:
            return False

        self.result = {
            'export_done': True,
            'filepath': filepath,
        }
        return True
