"""Download table C's original XLSX through visible WPS export controls."""
import datetime as dt
import os
from pathlib import Path
import tempfile
import time

from online_format import OnlineFormatter

EXCEL_LABELS = ('导出为 Excel(.xlsx)', '导出为Excel(.xlsx)',
                '导出为 Excel (.xlsx)', '导出为Excel（.xlsx）',
                'Excel (.xlsx)', 'Excel（.xlsx）', 'Excel文件(.xlsx)',
                'Microsoft Excel (.xlsx)', 'Excel', '.xlsx', 'XLSX',
                '导出为Excel', '导出为Excel文件', '另存为Excel文件', '下载为Excel')
EXPORT_LABELS = ('导出为', '下载为', '导出', '下载到本地', '下载', '另存为')


def _main_menu(ui):
    for label in ('文件', 'File', '文件菜单', '主菜单', '菜单', 'Menu'):
        control = _control(ui, label)
        if control is not None:
            return control
    # The compact WPS skin has an unlabelled hamburger instead of "文件".
    # Require menu semantics, a compact top-left location, and a unique target;
    # never click a hard-coded screen coordinate or an arbitrary toolbar icon.
    found = []
    for frame in ui._frames():
        paths = frame.locator('[class*="menu" i],[class*="hamburger" i],'
                              '[aria-haspopup="menu"]').evaluate_all(r'''elements => {
            const targets = new Set();
            for (const el of elements) {
                const node = el.closest('button,[role="button"],[aria-haspopup="menu"]') || el;
                if (node.closest('[role="grid"],[role="gridcell"],td,th,[role="dialog"],[role="menu"]')) continue;
                const box = node.getBoundingClientRect(), style = getComputedStyle(node);
                if (style.visibility !== 'visible' || style.display === 'none' ||
                    box.x < 0 || box.x > 100 || box.y < 0 || box.y > 100 ||
                    box.width < 12 || box.width > 64 || box.height < 12 || box.height > 64) continue;
                if (!['', '☰', '≡'].includes((node.innerText || '').trim())) continue;
                targets.add(node);
            }
            // A menu wrapper and its nested icon describe one target.
            return [...targets].filter(node => ![...targets].some(other => other !== node && other.contains(node)))
                .map(node => {
                    const parts = [];
                    for (let el = node; el && el.nodeType === 1; el = el.parentElement) {
                        let i = 1;
                        for (let prev = el.previousElementSibling; prev; prev = prev.previousElementSibling)
                            if (prev.tagName === el.tagName) i++;
                        parts.unshift(el.tagName.toLowerCase() + ':nth-of-type(' + i + ')');
                    }
                    return parts.join(' > ');
                });
        }''')
        found.extend(frame.locator(path) for path in paths)
    if len(found) > 1:
        raise ValueError('表C左上角文件菜单不唯一，已停止下载')
    if found and not found[0].is_enabled():
        raise ValueError('表C左上角文件菜单不可用，请检查权限')
    return found[0] if found else None


def _expand_export(ui, control, label):
    # WPS's "导出为" is a hover submenu. Clicking again can close it.
    control.hover(timeout=3000)
    ui.page.wait_for_timeout(250)
    if any(_control(ui, excel) is not None for excel in EXCEL_LABELS):
        return
    # Retain support for skins whose submenu opens only on click.
    ui._click(control, label)


def _control(ui, label, scope=None):
    item = ui._unique_control((label,), scope=scope, optional=True)
    if item is not None:
        return item
    # Some WPS menus render div/span labels without ARIA roles. Only accept
    # exact text inside a visible menu/control, never a spreadsheet cell.
    found = []
    for root in ([scope] if scope is not None else ui._frames()):
        for candidate in root.get_by_text(label, exact=True).all():
            if candidate.is_visible() and candidate.evaluate('''el =>
                !el.closest('[role="grid"],[role="gridcell"],td,th') &&
                !!el.closest('button,[role="button"],[role="menu"],[role="menuitem"],'
                    + '[class*="menu" i],[role="dialog"]')'''):
                found.append(candidate)
    if len(found) > 1:
        raise ValueError('表C下载菜单中“' + label + '”不唯一，已停止下载')
    if found:
        if not found[0].is_enabled():
            raise ValueError('表C的“' + label + '”不可用，请检查下载权限')
        return found[0]
    return None


def _dismiss_freeze_notice(ui):
    # This informational overlay blocks the menu. Ignore it without changing
    # the workbook's frozen panes or any other dialog/settings.
    message = '当前工作表的冻结线在显示区域之外，无法滚动浏览。若要浏览全部内容，请临时取消冻结窗格，此操作不会影响其他人。'
    for frame in ui._frames():
        notice = frame.get_by_text(message, exact=True)
        ignore = frame.get_by_role('button', name='忽略', exact=True)
        if notice.count() == 1 and notice.is_visible() and ignore.count() == 1 and ignore.is_visible():
            ignore.click(timeout=3000)
            notice.wait_for(state='hidden', timeout=3000)
            return True
    return False


def request_download(page, progress, should_stop=lambda: False, timeout=120, source_label='表C', diagnostic_dir=None):
    """Use normal export menus and accept only a real browser Download event."""
    ui = OnlineFormatter(page, lambda ref: None, progress=progress, diagnostic_dir=diagnostic_dir)
    downloads = []
    listener = lambda download: downloads.append(download)
    page.on('download', listener)
    try:
        opened = exported = selected = confirmed = False
        deadline = time.monotonic() + timeout
        while not downloads:
            if should_stop():
                from sync_runner import SyncStopped
                raise SyncStopped()
            if time.monotonic() >= deadline:
                exc = ValueError(f'{source_label}下载超时：未收到Excel文件，请检查{source_label}登录状态、下载权限和导出菜单')
                exc.diagnostic_context = {'visible_controls': ui._control_inventory(), 'screenshot': ui._capture_ui_dump(),
                                          'opened': opened, 'exported': exported, 'selected': selected, 'confirmed': confirmed}
                raise exc
            _dismiss_freeze_notice(ui)
            # Explicit format choices take precedence over generic export menus.
            control = None
            expand = False
            if not selected:
                for label in EXCEL_LABELS:
                    control = _control(ui, label)
                    if control is not None:
                        selected = True
                        break
            if control is None and not exported and not selected:
                for label in EXPORT_LABELS:
                    control = _control(ui, label)
                    if control is not None:
                        exported = True
                        expand = label in ('导出为', '下载为', '导出', '另存为')
                        break
            if control is None and not opened and not exported and not selected:
                control = _main_menu(ui)
                if control is not None:
                    opened = True
                    label = '文件菜单'
            if control is None and (selected or exported) and not confirmed:
                # Never click a generic confirmation outside an export dialog.
                dialogs = ui._visible_dialogs()
                if len(dialogs) == 1:
                    for label in ('下载', '导出', '确定'):
                        control = _control(ui, label, scope=dialogs[0])
                        if control is not None:
                            confirmed = True
                            break
            if control is not None:
                progress(f'正在下载{source_label}：' + label + '…')
                if expand:
                    _expand_export(ui, control, label)
                else:
                    try:
                        ui._click(control, label)
                    except ValueError:
                        if not _dismiss_freeze_notice(ui):
                            raise
                        # The notice can arrive between locating and clicking.
                        opened = exported = selected = confirmed = False
            else:
                page.wait_for_timeout(250)
        return downloads[0]
    finally:
        page.remove_listener('download', listener)


def save_download(download, root, sheet_name, *, target_name='表C.xlsx', label='表C'):
    """Validate the downloaded workbook before atomically replacing the target file."""
    from format_reference import load
    folder = Path(root) / 'DATA'
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / target_name
    fd, filename = tempfile.mkstemp(prefix='.table_download_', suffix='.xlsx', dir=folder)
    os.close(fd)
    temporary = Path(filename)
    try:
        failure = download.failure()
        if failure:
            raise ValueError(f'{label}下载失败：' + str(failure))
        if Path(download.suggested_filename).suffix.lower() != '.xlsx':
            raise ValueError(f'{label}下载结果不是.xlsx文件，请在{label}导出菜单选择Excel格式')
        download.save_as(str(temporary))
        try:
            reference = load(root, {'source': {'format_file': str(temporary.resolve()),
                                              'format_sheet': sheet_name}})
        except Exception as exc:
            raise ValueError(f'下载的{label}不是有效的Excel工作簿或缺少指定工作表，原{label}保持不变：' + str(exc)) from exc
        try:
            os.replace(temporary, target)
        except PermissionError as exc:
            raise ValueError(f'无法替换 DATA/{target_name}，请关闭正在打开{label}的Excel/WPS后重试') from exc
        return dict(path=f'DATA/{target_name}', sheet=reference['sheet'],
                    size=target.stat().st_size,
                    message=f'{label}已下载到程序根目录 DATA/{target_name}。')
    finally:
        temporary.unlink(missing_ok=True)


def download_table_b(copier, config, page=None):
    """下载表A（线上目标表）为 DATA/表B.xlsx；下载方法完全照搬表C。

    表B 是表A 的原始 xlsx 快照（数据 + 格式），用于在本地与表C 对比找出差异单元格，
    再只把差异粘贴回线上表A。
    """
    from sync_runner import _stop_requested
    from local_mirror import mirror_path
    if mirror_path(copier.root, config).resolve() == (Path(copier.root) / 'DATA' / '表B.xlsx').resolve():
        raise ValueError('表A1路径不能使用 DATA/表B.xlsx，请为镜像选择另一个文件')
    if page is None:
        page = copier._page_a(config['upload']['share_url'])
    progress = lambda message: copier.update_upload(message=message)
    progress('正在从表A下载原始Excel文件为表B…')
    download = request_download(page, progress, lambda: _stop_requested(copier), source_label='表A')
    progress('正在保存并校验 DATA/表B.xlsx…')
    return save_download(download, copier.root, config['upload']['sheet_name'],
                         target_name='表B.xlsx', label='表B')


def download_table_c(copier, config, page=None):
    from sync_runner import _stop_requested
    from local_mirror import mirror_path
    if mirror_path(copier.root, config).resolve() == (Path(copier.root) / 'DATA' / '表C.xlsx').resolve():
        raise ValueError('表A1路径不能使用 DATA/表C.xlsx，请为镜像选择另一个文件')
    if page is None:
        # This exports the whole workbook, so selecting a worksheet or reading
        # the canvas/namebox is both unnecessary and unreliable in WPS skins.
        page = copier._page_c(config['source']['share_url'])
    progress = lambda message: copier.update_upload(message=message)
    progress('正在从表C下载原始Excel文件为表C…')
    download = request_download(page, progress, lambda: _stop_requested(copier), source_label='表C',
                                diagnostic_dir=Path(copier.root) / 'error_reports')
    progress('正在保存并校验 DATA/表C.xlsx…')
    return save_download(download, copier.root, config['source']['sheet_name'],
                         target_name='表C.xlsx', label='表C')


# ---- 表B 历史快照（表A 整表快照）----
HSTY_DIR = 'DATA/HSTY'
SNAPSHOT_KEEP = 15  # DATA/HSTY 目录最多保留的快照数，超出删除最早的一个


def _snapshot_folder(root):
    folder = Path(root) / HSTY_DIR
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def snapshot_path(root):
    """生成带时间戳的快照路径 DATA/HSTY/表A_YYYYMMDD_HHMMSS.xlsx。"""
    stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
    return _snapshot_folder(root) / f'表A_{stamp}.xlsx'


def list_snapshots(root):
    """按修改时间从新到旧返回 DATA/HSTY 下的表A快照文件。"""
    folder = Path(root) / HSTY_DIR
    if not folder.is_dir():
        return []
    available = []
    for path in folder.glob('表A_*.xlsx'):
        try:
            available.append((path.stat().st_mtime, path.name, path))
        except FileNotFoundError:
            # The upload worker can rotate snapshots during a status request.
            continue
    return [item[2] for item in sorted(available, reverse=True)]


def prune_snapshots(root, keep=SNAPSHOT_KEEP):
    """超出 keep 份时删除最早的快照，返回本次删除的数量。"""
    snapshots = list_snapshots(root)
    removed = 0
    for old in snapshots[keep:]:
        try:
            old.unlink(missing_ok=True)
            removed += 1
        except OSError:
            pass
    return removed


def save_snapshot(download, root, sheet_name, *, label='表B'):
    """校验下载的表A，另存为带时间戳的历史快照到 DATA/HSTY，并轮换到最多 15 份。"""
    from format_reference import load
    target = snapshot_path(root)
    fd, filename = tempfile.mkstemp(prefix='.table_download_', suffix='.xlsx', dir=target.parent)
    os.close(fd)
    temporary = Path(filename)
    try:
        failure = download.failure()
        if failure:
            raise ValueError(f'{label}下载失败：' + str(failure))
        if Path(download.suggested_filename).suffix.lower() != '.xlsx':
            raise ValueError(f'{label}下载结果不是.xlsx文件，请在表A导出菜单选择Excel格式')
        download.save_as(str(temporary))
        try:
            load(root, {'source': {'format_file': str(temporary.resolve()),
                                   'format_sheet': sheet_name}})
        except Exception as exc:
            raise ValueError(f'下载的{label}不是有效的Excel工作簿或缺少指定工作表：' + str(exc)) from exc
        try:
            os.replace(temporary, target)
        except PermissionError as exc:
            raise ValueError(f'无法保存快照到 {target}，请关闭正在打开该文件的Excel/WPS后重试') from exc
        removed = prune_snapshots(root)
        return dict(path=target.relative_to(root).as_posix(), sheet=sheet_name,
                    size=target.stat().st_size, removed=removed,
                    message=f'{label}快照已保存到 DATA/HSTY/{target.name}。')
    finally:
        temporary.unlink(missing_ok=True)


def download_table_b_snapshot(copier, config):
    """下载表A为历史快照（DATA/HSTY，最多保留 15 份）。"""
    from sync_runner import _stop_requested
    activity = getattr(copier, 'activity', None)
    if activity is not None:
        activity.record('开始下载表B快照', path=HSTY_DIR)
    page = copier._page_a(config['upload']['share_url'])
    progress = lambda message: copier.update_upload(message=message)
    progress('正在下载表A为历史快照…')
    download = request_download(page, progress, lambda: _stop_requested(copier), source_label='表A')
    progress('正在保存并校验快照…')
    return save_snapshot(download, copier.root, config['upload']['sheet_name'], label='表B')
