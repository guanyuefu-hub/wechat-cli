import importlib.util
import sys
from pathlib import Path


MODULE_PATH = Path(r"C:\Users\guany\wechat-cli\.omx\prewarm_wechat_group.py")


def load_module():
    spec = importlib.util.spec_from_file_location("prewarm_wechat_group", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def test_build_key_image_points_returns_requested_count():
    module = load_module()
    rect = module.WindowRect(left=0, top=0, right=1000, bottom=1000)

    points = module.build_key_image_points(rect, 4)

    assert len(points) == 4
    assert points[0] == (680, 300)
    assert points[-1] == (680, 660)


def test_open_group_relies_on_keyboard_confirm_instead_of_clicking_fixed_session(monkeypatch):
    module = load_module()
    rect = module.WindowRect(left=0, top=0, right=1000, bottom=1000)
    clicks = []
    hotkeys = []
    presses = []

    monkeypatch.setattr(module, "click", lambda point: clicks.append(point))
    monkeypatch.setattr(module, "hotkey", lambda *keys: hotkeys.append(keys))
    monkeypatch.setattr(module, "press", lambda key: presses.append(key))
    monkeypatch.setattr(module, "assert_target_group_opened", lambda rect, group_name: None)
    monkeypatch.setattr(module.pyperclip, "copy", lambda text: None)
    monkeypatch.setattr(module.time, "sleep", lambda _: None)

    module.open_group(rect, "关富韵达中通客服群（220012）")

    assert clicks == [rect.point(0.115, 0.08)]
    assert hotkeys[:2] == [
        (module.win32con.VK_CONTROL, 0x41),
        (module.win32con.VK_CONTROL, 0x56),
    ]
    assert presses.count(module.win32con.VK_RETURN) == 2


def test_group_name_matches_ocr_text_tolerates_spaces_and_noise():
    module = load_module()

    matched = module.group_name_matches_ocr_text(
        "关富韵达中通客服群（220012）",
        "O 关 富 韵 达 中 通 客 服 群 （ 220012 〕 21 ） 0",
    )

    assert matched is True


def test_open_group_raises_when_target_group_not_confirmed(monkeypatch):
    module = load_module()
    rect = module.WindowRect(left=0, top=0, right=1000, bottom=1000)

    monkeypatch.setattr(module, "click", lambda point: None)
    monkeypatch.setattr(module, "hotkey", lambda *keys: None)
    monkeypatch.setattr(module, "press", lambda key: None)
    monkeypatch.setattr(module.pyperclip, "copy", lambda text: None)
    monkeypatch.setattr(module.time, "sleep", lambda _: None)
    monkeypatch.setattr(
        module,
        "read_current_chat_title_text",
        lambda rect: "腾讯新闻",
    )

    try:
        module.open_group(rect, "关富韵达中通客服群（220012）")
    except RuntimeError as exc:
        assert "未确认已切到目标群" in str(exc)
        assert "腾讯新闻" in str(exc)
    else:
        raise AssertionError("expected RuntimeError when OCR validation does not match")
