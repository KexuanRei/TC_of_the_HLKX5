"""Replace an existing cloud file through the WPS personal-center upload UI.

No verification: the frozen values workbook is uploaded under the same filename
and the overwrite is confirmed directly. The upload-complete notice ends the
round; nothing is downloaded or compared afterwards.
"""
import datetime as dt
import json
from pathlib import Path
import re
import time

from values_workbook import freeze_workbook

HOME = 'https://www.kdocs.cn/?show=all'


def validate_filename(value):
    if not isinstance(value, str):
        raise ValueError('云端文件名必须为文本')
    value = value.strip()
    if value and (not value.lower().endswith('.xlsx') or
                  re.search(r'[<>:"/\\|?*\x00-\x1f]', value) or
                  value.rstrip(' .') != value or value.startswith('.')):
        raise ValueError('云端文件名须为完整的.xlsx文件名，不能包含路径或特殊字符')
    return value


class CloudFiles:
    def __init__(self, page, stop=lambda: False, progress=lambda text: None):
        self.page, self.stop, self.progress = page, stop, progress

    def check(self):
        if self.stop():
            from sync_runner import SyncStopped
            raise SyncStopped()

    def open(self):
        self.page.goto(HOME, wait_until='domcontentloaded', timeout=60000)
        self.page.get_by_test_id('myDocument').wait_for(state='visible', timeout=60000)

    def row(self, filename):
        rows = self.page.locator('.item-container').filter(
            has=self.page.locator('[action-key="action_share"]'))
        # Attribute filtering, not the visually truncated/reversed filename.
        matches = [row for row in rows.all() if row.get_attribute('aria-label') == filename and row.is_visible()]
        if len(matches) != 1:
            raise ValueError(f'“我的云文档”根目录中未找到唯一的 {filename}。请确认原表A就在该目录；不会新建文件或覆盖其他目录。')
        return matches[0]

    def upload(self, path, filename):
        self.check()
        self.row(filename)
        # File input is used only after the visible import/upload action. WPS
        # renders two Import buttons on the root page; use the sidebar one.
        imports = self.page.locator('#uploaderBtn button').filter(has_text=re.compile(r'^\s*导入\s*$'))
        imports.first.click()
        with self.page.expect_file_chooser(timeout=15000) as chooser:
            self.page.get_by_text('上传文件', exact=True).and_(self.page.locator(':visible')).click()
        self.check()
        try:
            chooser.value.set_files(str(path))
            title = self.page.get_by_text('覆盖文件', exact=True)
            title.wait_for(state='visible', timeout=30000)
            # Find the unique visible modal with an overwrite button and confirm
            # directly; never press "保留两者".
            modal = self.page.locator('[role="dialog"],.kdv-dialog').filter(has=title)
            candidates = [item for item in modal.all() if item.is_visible() and item.get_by_role('button', name='覆盖', exact=True).count() == 1]
            if len(candidates) != 1:
                # Some WPS versions use an unlabelled wrapper around the title.
                modal = title.locator('xpath=..')
                for _ in range(4):
                    if modal.get_by_role('button', name='覆盖', exact=True).count() == 1:
                        break
                    modal = modal.locator('xpath=..')
            else:
                modal = candidates[0]
            self.check()
            modal.get_by_role('button', name='覆盖', exact=True).click()
            title.wait_for(state='hidden', timeout=30000)
            self.wait_upload()
        except Exception:
            cancel = self.page.get_by_role('button', name='取消', exact=True)
            if cancel.count() == 1 and cancel.is_visible():
                cancel.click()
            raise

    def wait_upload(self):
        """Do not navigate away while WPS is still sending the file."""
        deadline = time.monotonic() + 900
        while time.monotonic() < deadline:
            self.check()
            text = self.page.locator('.file-notify').inner_text()
            if re.search(r'上传失败|网络异常|上传出错', text):
                raise ValueError('WPS上传失败；已保留本地上传文件，下一轮将重新上传覆盖')
            if re.search(r'上传完成|上传成功', text) and not re.search(r'正在上传|等待上传', text):
                return
            self.progress('云文件上传中，请勿关闭专用Edge；上传完成后本轮同步即完成')
            self.page.wait_for_timeout(500)
        raise ValueError('云文件上传在15分钟内未完成，请在个人中心检查上传状态')


def _save_marker(path, record):
    path.write_text(json.dumps({'filename': record['filename']}, ensure_ascii=False, indent=2),
                    encoding='utf-8')


def _retain_uploads(root, record):
    """Retain three completed upload packages; never touch pending packages."""
    folder = Path(record['path']).resolve().parent
    storage = (Path(root) / 'DATA' / 'UPLOAD').resolve()
    if not folder.is_relative_to(storage):
        return
    _save_marker(folder / 'completed.json', record)
    completed = sorted(storage.glob('*/completed.json'), key=lambda p: p.stat().st_mtime, reverse=True)
    for marker in completed[3:]:
        old = marker.parent.resolve()
        if not old.is_relative_to(storage) or not re.fullmatch(r'\d{8}_\d{6}_\d{6}', old.name):
            continue
        try:
            filename = validate_filename(json.loads(marker.read_text(encoding='utf-8'))['filename'])
            if not filename or {p.name for p in old.iterdir()} - {filename, 'completed.json'}:
                continue
            for name in (filename, 'completed.json'):
                (old / name).unlink(missing_ok=True)
            old.rmdir()
        except (OSError, ValueError, KeyError):
            pass


def replace_cloud_file(copier, config):
    from sync_runner import _stop_requested
    root = Path(copier.root)
    stop = lambda: _stop_requested(copier)
    progress = lambda message: copier.update_upload(message=message)
    copier._ensure_context()
    page = copier.context.new_page()
    ui = CloudFiles(page, stop, progress)
    try:
        filename = validate_filename(config['upload'].get('cloud_filename', ''))
        if not filename:
            page_a = copier._page_a(config['upload']['share_url'])
            filename = validate_filename(page_a.title().strip() + ('' if page_a.title().lower().endswith('.xlsx') else '.xlsx'))
        if not filename:
            raise ValueError('请填写表A在个人中心显示的完整云端文件名')
        source = Path(config['source'].get('format_file') or 'DATA/表C.xlsx')
        if not source.is_absolute():
            source = root / source
        stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        target = root / 'DATA' / 'UPLOAD' / stamp / filename
        progress('离线转换：保留表C所有工作表与格式，仅移除单元格公式')
        record = freeze_workbook(source, target, stop, progress)
        record['filename'] = filename
        ui.open()
        # Wait for the exact filename after the asynchronous file listing.
        ui.page.locator('.item-container').filter(has=ui.page.get_by_role('button', name='分享', exact=True)).first.wait_for(state='visible', timeout=60000)
        progress('上传纯值文件并直接覆盖原表A')
        ui.upload(target, filename)
        # Whole-file replacement also replaces worksheet names. Keep the
        # existing snapshot/open-sheet actions pointed at the source sheet.
        updated = json.loads(json.dumps(config))
        updated['upload']['cloud_filename'] = record['filename']
        updated['upload']['sheet_name'] = config['source']['sheet_name']
        copier._write_config_file(updated)
        copier.set_sync_config(updated)
        _retain_uploads(root, record)
        from table_c_download import prune_snapshots
        prune_snapshots(root)
        message = f'原表A文件上传覆盖完成。已去除 {record["formulas"]} 个公式，保留全部工作表和格式。'
        if record['formula_errors']:
            message += f' 表C原有 {record["formula_errors"]} 个公式错误值已原样保留。'
        return dict(uploaded=record['cells'], formulas=record['formulas'],
                    formula_errors=record['formula_errors'], path=record['path'], message=message)
    finally:
        page.close()
