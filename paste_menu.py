"""Dismiss WPS paste options without applying another paste mode."""
from online_format import OnlineFormatter


def select_paste_mode(page, mode):
    """Choose the explicit WPS paste option; never silently use normal paste."""
    labels = {'formats': '粘贴格式', 'values': '粘贴值'}
    if mode not in labels:
        raise ValueError('未知粘贴模式')
    label = labels[mode]
    formatter = OnlineFormatter(page, lambda _: None)
    opened = False
    for _ in range(20):
        candidates = []
        for frame in formatter._frames():
            for item in frame.get_by_text(label, exact=True).all():
                if item.is_visible() and item.evaluate('''el => {
                    if(el.closest('[role="grid"],[role="gridcell"],td,th'))return false;
                    if(el.closest('[aria-disabled="true"],[disabled]'))return false;
                    for(let p=el;p && p!==document.body;p=p.parentElement){
                        const text=p.innerText||'';
                        if(p.matches('[role="menu"],[role="listbox"],[class*="menu" i],[class*="popup" i],[class*="dropdown" i]') &&
                           text.includes('粘贴值') && text.includes('粘贴格式'))return true;
                    }
                    return false;
                }'''):
                    candidates.append(item)
        if len(candidates) > 1:
            raise ValueError(f'表A存在多个“{label}”菜单项，未选择粘贴模式')
        if candidates:
            candidates[0].click()
            page.wait_for_timeout(300)
            dismiss_paste_options(page)
            return
        if not opened:
            controls = formatter._controls(['粘贴选项'])
            if len(controls) > 1:
                raise ValueError('表A存在多个粘贴选项按钮，未选择粘贴模式')
            if controls:
                controls[0][0].click()
                opened = True
        page.wait_for_timeout(150)
    raise ValueError(f'表A未出现“{label}”菜单，不能确认粘贴模式，已停止本轮')


def dismiss_paste_options(page):
    """Only Escape a visible paste menu, not arbitrary dialogs or cell text."""
    def visible():
        for frame in OnlineFormatter(page, lambda _: None)._frames():
            for item in frame.get_by_text('粘贴值和数字格式', exact=True).all():
                if item.is_visible() and item.evaluate('''el => {
                    if (el.closest('[role="grid"],[role="gridcell"],td,th')) return false;
                    for (let node = el; node && node !== document.body; node = node.parentElement) {
                        const text = node.innerText || '';
                        if (node.matches('[role="menu"],[role="listbox"],[class*="menu" i],[class*="popup" i],[class*="dropdown" i]') &&
                            text.includes('粘贴值') && text.includes('粘贴格式')) return true;
                    }
                    return false;
                }'''):
                    return True
        return False
    for _ in range(2):
        if not visible():
            return
        page.keyboard.press('Escape')
        page.wait_for_timeout(100)
    if visible():
        raise ValueError('表A粘贴选项菜单未关闭，已停止后续操作；请关闭菜单后重试')
