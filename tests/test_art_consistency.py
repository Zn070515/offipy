from art_helpers import make_scene, make_slide, make_text_element
from offipy.art.consistency import RULE_MARGIN_DRIFT, RULE_TITLE_DRIFT, assess_deck
from offipy.art.features import infer_slide_role
from offipy.art.profiles import get_profile


def _content_slide(index, title_x=0.1, title_size=48.0, body_x=0.1, **kw):
    return make_slide(
        index=index,
        elements=[
            make_text_element(
                "t", "Title", x=title_x, y=0.05, w=0.4, h=0.08, font_size=title_size, role="title"
            ),
            make_text_element(
                "b", "Body", x=body_x, y=0.2, w=0.5, h=0.06, font_size=24.0, role="body"
            ),
            make_text_element(
                "b2", "Body2", x=body_x, y=0.3, w=0.5, h=0.06, font_size=20.0, role="body"
            ),
        ],
    )


def test_infer_role_groups_content():
    scene = make_scene([_content_slide(1), _content_slide(2), _content_slide(3)])
    roles = {infer_slide_role(s) for s in scene.slides}
    assert roles == {"content"}  # ≥3 元素 → content，同组


def test_title_drift_across_slides():
    scene = make_scene(
        [
            _content_slide(1, title_x=0.1),
            _content_slide(2, title_x=0.5),
            _content_slide(3, title_x=0.1),
        ]
    )
    assert any(f.rule_id == RULE_TITLE_DRIFT for f in assess_deck(scene, get_profile("balanced")))


def test_no_title_drift():
    scene = make_scene(
        [
            _content_slide(1, title_x=0.1),
            _content_slide(2, title_x=0.1),
            _content_slide(3, title_x=0.1),
        ]
    )
    assert all(f.rule_id != RULE_TITLE_DRIFT for f in assess_deck(scene, get_profile("balanced")))


def test_insufficient_slides_no_drift():
    scene = make_scene(
        [
            _content_slide(1, title_x=0.1),
            _content_slide(2, title_x=0.5),
        ]
    )
    assert assess_deck(scene, get_profile("balanced")) == []


def test_margin_drift_across_slides():
    # 第三页标题和正文都移到 x=0.6 → 最小左边距从 0.1 变 0.6 → 漂移
    scene = make_scene(
        [
            _content_slide(1, body_x=0.1),
            _content_slide(2, body_x=0.1),
            _content_slide(3, title_x=0.6, body_x=0.6),
        ]
    )
    assert any(f.rule_id == RULE_MARGIN_DRIFT for f in assess_deck(scene, get_profile("balanced")))


def test_margin_single_body_stays_min_no_drift():
    # 第三页只有一个 body 移到 x=0.6、另一个仍在 x=0.1 → min 仍 0.1，不误报
    slide3 = make_slide(
        3,
        elements=[
            make_text_element(
                "t", "Title", x=0.1, y=0.05, w=0.4, h=0.08, font_size=48.0, role="title"
            ),
            make_text_element(
                "b", "Body", x=0.6, y=0.2, w=0.5, h=0.06, font_size=24.0, role="body"
            ),
            make_text_element(
                "b2", "Body2", x=0.1, y=0.3, w=0.5, h=0.06, font_size=20.0, role="body"
            ),
        ],
    )
    scene = make_scene([_content_slide(1, body_x=0.1), _content_slide(2, body_x=0.1), slide3])
    assert all(f.rule_id != RULE_MARGIN_DRIFT for f in assess_deck(scene, get_profile("balanced")))


def test_filters_decoration_from_margin():
    from offipy.art.models import ArtElement

    deco = ArtElement(
        element_id="d",
        kind="shape",
        role="decoration",
        x=0.0,
        y=0.9,
        width=1.0,
        height=0.05,
        slide_index=1,
        decoration=True,
    )

    def sl(i):
        return make_slide(
            i,
            elements=[
                make_text_element("a", "A", x=0.1, font_size=24.0),
                make_text_element("b", "B", x=0.1, y=0.2, font_size=20.0),
                deco,
            ],
        )

    scene = make_scene([sl(1), sl(2), sl(3)])
    assert all(f.rule_id != RULE_MARGIN_DRIFT for f in assess_deck(scene, get_profile("balanced")))


def test_title_size_drift_detected():
    scene = make_scene(
        [
            _content_slide(1, title_x=0.1, title_size=48.0),
            _content_slide(2, title_x=0.1, title_size=48.0),
            _content_slide(3, title_x=0.1, title_size=30.0),
        ]
    )
    assert any(f.rule_id == RULE_TITLE_DRIFT for f in assess_deck(scene, get_profile("balanced")))


def test_different_roles_do_not_cross_compare():
    # 2 个 cover + 3 个 content：cover 不足 3 不判；content 同组才判
    covers = [
        make_slide(
            i,
            elements=[
                make_text_element("t", "Cover", x=0.1, y=0.05, font_size=52.0, role="title"),
                make_text_element("s", "Sub", x=0.1, y=0.3, font_size=26.0, role="subtitle"),
            ],
        )
        for i in (1, 2)
    ]
    contents = [_content_slide(i) for i in (3, 4, 5)]
    scene = make_scene(covers + contents)
    # content 组标题一致 → 无 title drift
    assert all(f.rule_id != RULE_TITLE_DRIFT for f in assess_deck(scene, get_profile("balanced")))
