"""
validate_plan.py
================

Pre-build validator for the Bmatix PPT skill.

Runs a sequence of HARD rule checks against a plan.json BEFORE the build
starts. Any violation raises PlanValidationError with a structured list of
issues, each carrying enough context for Claude to auto-fix and retry.

The validator is the source of truth for hard rules — what's enforced here
is what's enforced, regardless of what SKILL.md says. SKILL.md describes
intent; this file describes enforcement.

Rules enforced (HARD — build fails on violation):
  R1  No deprecated black-and-white Unicode glyphs (✓ ✗ ★ → ▶) in any icon field
  R2  No paragraph >40 words in any card body
  R3  cards_2_large card bodies must be broken (bullets or line breaks, no prose)
  R4  Lavender-pair rule: 4-lavender slide must be followed by a breaker slide
  R5  Lavender density per slide: max 4 same-shape, max 5 total
  R6  Filled-header colors: all navy except auto-detected Result cards (orange)
  R7  Icon consistency per card group: all emoji OR all numbers, no mix
  R8  Pair-diversity: ≥2 visual variables differ between adjacent slides
  R9  Pill present (decks with 4+ content slides)
  R10 Example/Voorbeeld card present (decks with 4+ content slides)
  R11 Per-layout word ceilings for card bodies and titles at 16pt body / 20pt title

Usage:
  from validate_plan import validate_plan, PlanValidationError

  try:
      validate_plan(plan)
  except PlanValidationError as e:
      for issue in e.issues:
          print(f"[{issue.rule}] slide {issue.slide_index}: {issue.message}")
"""

from dataclasses import dataclass, field
from typing import Optional

# Import the canonical auto-detection helpers from slide_helpers, so the
# validator and the renderer always agree on which cards count as Example
# or Result. Single source of truth = slide_helpers.py.
from slide_helpers import _is_result_card, _is_example_card


# =============================================================================
# DATA STRUCTURES
# =============================================================================

@dataclass
class ValidationIssue:
    """A single validation failure with enough context for auto-fix."""
    rule: str               # e.g. "R1" - the rule identifier
    slide_index: Optional[int]  # 1-based content-slide index, or None for deck-level
    message: str            # human-readable description of the problem
    fix_hint: str           # concrete suggestion for fixing the plan

    def __str__(self):
        loc = f"slide {self.slide_index}" if self.slide_index else "deck-level"
        return f"[{self.rule} | {loc}] {self.message}  →  fix: {self.fix_hint}"


class PlanValidationError(Exception):
    """Raised when plan.json fails one or more hard rules."""
    def __init__(self, issues: list):
        self.issues = issues
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        lines = [f"Plan validation failed with {len(self.issues)} issue(s):"]
        for issue in self.issues:
            lines.append(f"  {issue}")
        return "\n".join(lines)


# =============================================================================
# RULE CONSTANTS (kept in sync with slide_helpers.py auto-detection logic)
# =============================================================================

# Deprecated black-and-white Unicode glyphs — forbidden in any icon field
DEPRECATED_GLYPHS = {"✓", "✗", "★", "→", "▶", "●", "○", "■", "□"}

# Layouts that render cards as plain text (not auto-bulleted) — bodies in
# these layouts must be broken into bullets or line-separated phrases.
PROSE_PRONE_LAYOUTS = {"cards_2_large", "cards_2x2_compare"}

# Layouts where lavender card grids dominate visually — used by the
# lavender-pair rule (R4).
LAVENDER_GRID_LAYOUTS = {
    "cards_2x2_compare",
    "cards_2_large",
    "cards_numbered_steps",
    "card_1_hero_with_subcards",
    "cards_2_bulletlist_plus_wide",
}

# Layouts that BREAK the lavender pattern (satisfy lavender-pair rule)
LAVENDER_BREAKER_LAYOUTS = {
    "decision_tree",
    "cards_left_image_right",
    "card_hero_with_white_subcards",
}


# =============================================================================
# HELPER PREDICATES
# =============================================================================

# Note: _is_result_card and _is_example_card are imported from slide_helpers
# at the top of this file - that module is the canonical source of truth
# for these detections. Do not redefine them here.


def _is_image_card(card_spec: dict) -> bool:
    """Detect image placeholder cards."""
    return bool(card_spec.get("is_image"))


def _collect_cards(slide: dict) -> list:
    """Return ALL card-like dicts on a slide regardless of layout key.

    Card layouts use different keys (cards / steps / subcards / wide / branches).
    Hero cards are also included. `decision_tree` branches are skipped here
    because they're not lavender-card shapes (they're navy-banner step cards).
    """
    out = []
    for key in ("cards", "steps", "subcards"):
        items = slide.get(key, [])
        if isinstance(items, list):
            out.extend(items)
    for key in ("hero", "wide"):
        item = slide.get(key)
        if isinstance(item, dict):
            out.append(item)
    return out


def _count_lavender_cards(slide: dict) -> int:
    """Count lavender (default) cards on a slide. Beige, image, and result
    cards don't count as lavender; filled-header cards don't either.
    """
    if slide.get("header_style") == "filled":
        return 0  # filled headers turn cards into navy/orange bands

    layout = slide.get("layout", "")
    if layout in LAVENDER_BREAKER_LAYOUTS:
        return 0  # these layouts don't produce a flat lavender grid

    count = 0
    for card in _collect_cards(slide):
        if _is_image_card(card):
            continue  # image placeholders break the lavender pattern
        if _is_example_card(card):
            continue  # beige fill, not lavender
        # Result cards are STILL lavender (only stripe/title turn orange),
        # so they count.
        count += 1
    return count


def _slide_breaks_lavender_pattern(slide: dict) -> bool:
    """True if the slide is one of the lavender-pattern-breaking layouts
    OR uses a structural break (filled headers, image placeholder, example card).
    """
    layout = slide.get("layout", "")
    if layout in LAVENDER_BREAKER_LAYOUTS:
        return True
    if slide.get("header_style") == "filled":
        return True
    for card in _collect_cards(slide):
        if _is_image_card(card) or _is_example_card(card):
            return True
    return False


def _visual_signature(slide: dict) -> dict:
    """Compute the visual fingerprint of a slide for pair-diversity comparison.

    Each key in the returned dict is one "visual variable". Two slides differ
    on a variable iff their values for that key differ.
    """
    cards = _collect_cards(slide)
    return {
        "layout": slide.get("layout", ""),
        "header_style": slide.get("header_style", "external"),
        "has_pill": bool(slide.get("hero", {}).get("pill")) if isinstance(slide.get("hero"), dict) else False,
        "uses_emojis": any(c.get("icon") for c in cards if isinstance(c, dict)),
        "has_example": any(_is_example_card(c) for c in cards if isinstance(c, dict)),
        "has_result": any(_is_result_card(c) for c in cards if isinstance(c, dict)),
        "has_image": any(_is_image_card(c) for c in cards if isinstance(c, dict)),
        "uses_stripes": any(c.get("stripe") for c in cards if isinstance(c, dict)),
    }


def _count_words(text: str) -> int:
    """Word count, ignoring bullet/whitespace characters."""
    if not isinstance(text, str):
        return 0
    cleaned = text.replace("•", " ").replace("\n", " ")
    return len(cleaned.split())


def _is_prose_block(body: str) -> bool:
    """Detect prose-block style: long-ish text with no bullets and no line breaks.

    A body is prose-block if all of:
      - >25 words (short bodies are fine as a single sentence)
      - no '•' character
      - no '\n' character (no line breaks)
    """
    if not isinstance(body, str):
        return False
    if _count_words(body) <= 25:
        return False
    if "•" in body:
        return False
    if "\n" in body:
        return False
    return True


# =============================================================================
# INDIVIDUAL RULE CHECKS
# =============================================================================

def _check_r1_deprecated_glyphs(slides: list) -> list:
    """R1: No deprecated black-and-white Unicode glyphs (✓ ✗ ★ → ▶) in icon fields."""
    issues = []
    for i, slide in enumerate(slides, start=1):
        for card in _collect_cards(slide):
            icon = card.get("icon", "")
            if isinstance(icon, str):
                for glyph in DEPRECATED_GLYPHS:
                    if glyph in icon:
                        issues.append(ValidationIssue(
                            rule="R1",
                            slide_index=i,
                            message=f"deprecated glyph '{glyph}' in card icon",
                            fix_hint=f"replace '{glyph}' with a modern emoji (✅ ❌ ⭐ ➡️ ▶️ etc.)",
                        ))
    return issues


def _check_r2_paragraph_length(slides: list) -> list:
    """R2: No paragraph >40 words in any card body."""
    issues = []
    for i, slide in enumerate(slides, start=1):
        for card in _collect_cards(slide):
            body = card.get("body", "")
            if not isinstance(body, str):
                continue
            # Split on double newlines for paragraphs; also each single line counts
            paragraphs = [p.strip() for p in body.replace("\r", "").split("\n") if p.strip()]
            for para in paragraphs:
                wc = _count_words(para)
                if wc > 40:
                    label = card.get("label") or card.get("title") or "(no label)"
                    issues.append(ValidationIssue(
                        rule="R2",
                        slide_index=i,
                        message=f"card '{label}' has a {wc}-word paragraph (>40)",
                        fix_hint="split into shorter bullets, line-separated phrases, or split across two slides",
                    ))
    return issues


def _check_r3_cards_2_large_prose(slides: list) -> list:
    """R3: cards_2_large bodies must be broken into bullets/lines, not prose."""
    issues = []
    for i, slide in enumerate(slides, start=1):
        if slide.get("layout") != "cards_2_large":
            continue
        for card in slide.get("cards", []):
            if _is_image_card(card):
                continue  # image placeholders don't have body text
            body = card.get("body", "")
            if _is_prose_block(body):
                label = card.get("label") or card.get("title") or "(no label)"
                wc = _count_words(body)
                issues.append(ValidationIssue(
                    rule="R3",
                    slide_index=i,
                    message=f"card '{label}' in cards_2_large is a {wc}-word prose block (no bullets, no line breaks)",
                    fix_hint="rewrite body as bullets (each line starting with '•') or split into separate lines with '\\n'",
                ))
    return issues


def _check_r4_lavender_pair(slides: list) -> list:
    """R4: A slide with 4+ lavender cards must be followed by a breaker slide."""
    issues = []
    for i, slide in enumerate(slides, start=1):
        lav = _count_lavender_cards(slide)
        if lav >= 4 and i < len(slides):
            next_slide = slides[i]  # i is 1-based, so slides[i] is the next one
            if not _slide_breaks_lavender_pattern(next_slide):
                next_lav = _count_lavender_cards(next_slide)
                issues.append(ValidationIssue(
                    rule="R4",
                    slide_index=i + 1,  # the violating slide is the NEXT one
                    message=f"slide {i} has {lav} lavender cards but slide {i+1} doesn't break the pattern (has {next_lav} lavender cards, layout={next_slide.get('layout')})",
                    fix_hint=f"change slide {i+1} to one of: decision_tree, cards_left_image_right, card_hero_with_white_subcards, OR add header_style='filled', OR include an Example/Voorbeeld card, OR include is_image=true card",
                ))
    return issues


def _check_r5_lavender_density(slides: list) -> list:
    """R5: max 4 lavender cards of same shape, max 5 total per slide.

    'Same shape' here means same role in the layout — for current layouts this
    is enforced as 'total lavender card count'. The 5-total cap dominates; if
    a slide has 5 lavender cards they're typically the same shape anyway.
    """
    issues = []
    for i, slide in enumerate(slides, start=1):
        lav = _count_lavender_cards(slide)
        if lav > 5:
            issues.append(ValidationIssue(
                rule="R5",
                slide_index=i,
                message=f"slide has {lav} lavender cards (max 5)",
                fix_hint="swap one card to Example (beige), Result (orange stripe), set header_style='filled', or split into two slides",
            ))
        elif lav > 4:
            # 5 is allowed ONLY if the 5th is visually distinct (different shape/role)
            # For now we flag 5 as a soft warning by checking if the layout naturally
            # produces 5 same-shape cards. card_1_hero_with_subcards = 1 hero + 4 subs
            # is fine (hero is a different shape). cards_2x2_compare = 4 same-shape is
            # also fine (it's exactly 4). 5 same-shape is unusual.
            layout = slide.get("layout", "")
            if layout == "card_1_hero_with_subcards":
                continue  # hero + 4 subcards = 5 but hero is a distinct shape
            # Otherwise 5+ same-shape cards = violation
            issues.append(ValidationIssue(
                rule="R5",
                slide_index=i,
                message=f"slide has {lav} lavender cards of the same shape (max 4 same-shape)",
                fix_hint="swap one card to Example (beige), Result (orange stripe), or use header_style='filled' to break the pattern",
            ))
    return issues


def _check_r6_filled_header_colors(slides: list) -> list:
    """R6: filled-header bands all navy, except auto-detected Result cards (orange).

    The header color for filled headers is driven by `circle_color` on the
    individual card. Allowed values for filled-header cards:
      - navy (default)
      - orange ONLY if the card is auto-detected as a Result/Conclusion card

    Mixed green/red/orange filled headers on parallel cards = violation.
    """
    issues = []
    for i, slide in enumerate(slides, start=1):
        if slide.get("header_style") != "filled":
            continue
        for card in _collect_cards(slide):
            circle_color = card.get("circle_color", "navy")
            if circle_color == "navy":
                continue
            if circle_color == "orange":
                if _is_result_card(card):
                    continue  # allowed
                label = card.get("label") or card.get("title") or "(no label)"
                issues.append(ValidationIssue(
                    rule="R6",
                    slide_index=i,
                    message=f"card '{label}' has orange filled header but is not a Result/Conclusion card",
                    fix_hint="change circle_color to 'navy', OR rename the card to a result word (Result/Conclusion/Summary/Takeaway/Insights/etc.), OR set is_result=true",
                ))
            else:
                # green/red filled headers are always wrong
                label = card.get("label") or card.get("title") or "(no label)"
                issues.append(ValidationIssue(
                    rule="R6",
                    slide_index=i,
                    message=f"card '{label}' uses '{circle_color}' filled header (only navy or orange-on-Result allowed)",
                    fix_hint=f"change circle_color from '{circle_color}' to 'navy'; move status nuance to header_style='external' where mixed colors are allowed",
                ))
    return issues


def _check_r7_icon_consistency(slides: list) -> list:
    """R7: within one card group, all cards have emojis OR all show numbers.

    Applies to cards_numbered_steps (steps array) and card_1_hero_with_subcards
    (subcards array). No half-and-half.
    """
    issues = []
    for i, slide in enumerate(slides, start=1):
        layout = slide.get("layout", "")
        group_key = None
        if layout == "cards_numbered_steps":
            group_key = "steps"
        elif layout == "card_1_hero_with_subcards":
            group_key = "subcards"
        else:
            continue

        group = slide.get(group_key, [])
        if not isinstance(group, list) or len(group) == 0:
            continue

        icons_set = [bool(card.get("icon")) for card in group if isinstance(card, dict)]
        if any(icons_set) and not all(icons_set):
            missing = [
                idx + 1 for idx, has in enumerate(icons_set) if not has
            ]
            issues.append(ValidationIssue(
                rule="R7",
                slide_index=i,
                message=f"mixed numbers and emojis in {group_key} (cards at positions {missing} lack an icon while others have one)",
                fix_hint=f"either add a modern emoji to ALL {len(group)} cards, or remove icon from ALL cards (default to numbers)",
            ))
    return issues


def _check_r8_pair_diversity(slides: list) -> list:
    """R8: ≥2 visual variables differ between every adjacent pair."""
    issues = []
    if len(slides) < 2:
        return issues
    for i in range(len(slides) - 1):
        sig_a = _visual_signature(slides[i])
        sig_b = _visual_signature(slides[i + 1])
        differing = [k for k in sig_a if sig_a[k] != sig_b[k]]
        if len(differing) < 2:
            same = [k for k in sig_a if sig_a[k] == sig_b[k]]
            issues.append(ValidationIssue(
                rule="R8",
                slide_index=i + 2,  # report against the SECOND slide of the pair
                message=f"slides {i+1} and {i+2} only differ in {len(differing)} visual variable(s) (need ≥2); shared: {same}",
                fix_hint=f"vary at least 2 of: layout, header_style, has_pill, uses_emojis, has_example, has_result, has_image, uses_stripes",
            ))
    return issues


def _check_r9_pill_present(slides: list) -> list:
    """R9: at least 1 pill badge anywhere in decks with 4+ content slides.

    Threshold is 4 (not 3) to guarantee feasibility: at 1-3 slides the rule is
    skipped entirely, so a minimal deck is always valid and the auto-retry loop
    can never exhaust on a structurally impossible "somewhere in the deck"
    requirement. Kept in sync with R10 and with SKILL.md's Feasibility guarantee.
    """
    issues = []
    if len(slides) < 4:
        return issues
    for slide in slides:
        hero = slide.get("hero")
        if isinstance(hero, dict) and hero.get("pill"):
            return []  # found one
    issues.append(ValidationIssue(
        rule="R9",
        slide_index=None,
        message="no pill badge anywhere in deck (4+ content slides need ≥1 pill)",
        fix_hint="add a pill to one hero card: e.g. {'hero': {'pill': 'STATUS · IN PROGRESS', ...}}",
    ))
    return issues


def _check_r10_example_present(slides: list) -> list:
    """R10: at least 1 Example/Voorbeeld card anywhere in decks with 4+ content slides.

    Threshold is 4 (not 3) for the same feasibility reason as R9: below 4 slides
    the rule is skipped, never failed, so no valid slide count can produce an
    unsatisfiable plan. Kept in sync with R9 and SKILL.md's Feasibility guarantee.
    """
    issues = []
    if len(slides) < 4:
        return issues
    for slide in slides:
        for card in _collect_cards(slide):
            if _is_example_card(card):
                return []  # found one
    issues.append(ValidationIssue(
        rule="R10",
        slide_index=None,
        message="no Example/Voorbeeld card anywhere in deck (4+ content slides need ≥1)",
        fix_hint="add a card with 'Example: ...' or 'Voorbeeld: ...' in its label/title to ground the abstract content in something concrete",
    ))
    return issues


# Per-layout word ceilings for card BODY text at the current font sizes
# (FONT_CARD_BODY=16pt, FONT_CARD_TITLE=20pt). These are conservative
# heuristics tuned from visual QA of real decks. Going over the ceiling
# causes text to overflow visually.
#
# Format: layout_key → {role: max_words, ...}
# Roles: 'body' (default card body), 'hero_body' (wide hero card body),
# 'subcard_body' (smaller subcards beneath a hero).
LAYOUT_WORD_CEILINGS = {
    "cards_2_large":                  {"body": 65},    # 20pt body fills the large card with ~5-6 bullets
    "cards_2x2_compare":              {"body": 35},
    "cards_left_image_right":         {"body": 16},
    "cards_numbered_steps":           {"body": 20},   # tightened from visual QA
    "card_1_hero_with_subcards":      {"hero_body": 28, "subcard_body": 18},
    "card_hero_with_white_subcards":  {"hero_body": 15, "subcard_body": 14},   # tightened from visual QA
    "cards_2_bulletlist_plus_wide":   {"body": 45},
    "decision_tree":                  {"body": 6},
}

# Per-layout word ceilings for card TITLE / LABEL text at 20pt. Titles
# overflow visually if they wrap to too many lines.
LAYOUT_TITLE_CEILINGS = {
    "cards_2_large":                  5,
    "cards_2x2_compare":              4,
    "cards_left_image_right":         3,
    "cards_numbered_steps":           2,   # tightened — "Application submitted" already wraps
    "card_1_hero_with_subcards":      3,
    "card_hero_with_white_subcards":  3,
    "cards_2_bulletlist_plus_wide":   4,
    "decision_tree":                  3,
}


def _check_r11_overflow(slides: list) -> list:
    """R11: detect card body / title overflow based on per-layout word ceilings.

    The font sizes were bumped to 16pt body / 20pt title for readability at
    presentation distance. That means cards now hold fewer words before
    text spills over the card's bottom edge. This check catches overflow
    BEFORE the build, so Claude can trim the offending card and retry.
    """
    issues = []
    for i, slide in enumerate(slides, start=1):
        layout = slide.get("layout", "")
        body_ceilings = LAYOUT_WORD_CEILINGS.get(layout, {})
        title_ceiling = LAYOUT_TITLE_CEILINGS.get(layout, 6)

        # cards_numbered_steps gets tighter with more steps (cards are narrower).
        # 3 steps → 3-word title ceiling; 4 steps → 2-word title ceiling.
        if layout == "cards_numbered_steps":
            step_count = len(slide.get("steps", []))
            if step_count >= 4:
                title_ceiling = 2
                # Body ceiling also tightens with 4 steps
                body_ceilings = {"body": 18}
            elif step_count == 3:
                title_ceiling = 3
                body_ceilings = {"body": 22}
            else:  # 2 steps - wider cards
                title_ceiling = 4
                body_ceilings = {"body": 30}

        # Determine which role each card slot plays in this layout
        if layout == "card_1_hero_with_subcards" or layout == "card_hero_with_white_subcards":
            hero = slide.get("hero")
            if isinstance(hero, dict):
                hb = _count_words(hero.get("body", ""))
                ceiling = body_ceilings.get("hero_body", 40)
                if hb > ceiling:
                    issues.append(ValidationIssue(
                        rule="R11",
                        slide_index=i,
                        message=f"hero body has {hb} words (max {ceiling} at 16pt for {layout})",
                        fix_hint=f"trim hero body to ≤{ceiling} words — shorten the framing sentence to its core idea",
                    ))
            sub_ceiling = body_ceilings.get("subcard_body", 22)
            for j, card in enumerate(slide.get("subcards", []), start=1):
                if not isinstance(card, dict):
                    continue
                wb = _count_words(card.get("body", ""))
                if wb > sub_ceiling:
                    label = card.get("title") or card.get("label") or f"subcard {j}"
                    issues.append(ValidationIssue(
                        rule="R11",
                        slide_index=i,
                        message=f"subcard '{label}' has {wb}-word body (max {sub_ceiling} at 16pt for {layout})",
                        fix_hint=f"trim subcard '{label}' body to ≤{sub_ceiling} words",
                    ))
                # Title check
                tw = _count_words(card.get("title", ""))
                if tw > title_ceiling:
                    issues.append(ValidationIssue(
                        rule="R11",
                        slide_index=i,
                        message=f"subcard title '{card.get('title')}' has {tw} words (max {title_ceiling} at 20pt for {layout})",
                        fix_hint=f"shorten title to ≤{title_ceiling} words",
                    ))
        else:
            # Other layouts: check every card / step / cards array
            ceiling = body_ceilings.get("body", 30)
            for j, card in enumerate(_collect_cards(slide), start=1):
                if not isinstance(card, dict):
                    continue
                if _is_image_card(card):
                    continue  # image placeholders have no body to overflow
                # Body check
                body_field = card.get("body", "")
                if isinstance(body_field, str):
                    wb = _count_words(body_field)
                    if wb > ceiling:
                        label = card.get("label") or card.get("title") or f"card {j}"
                        issues.append(ValidationIssue(
                            rule="R11",
                            slide_index=i,
                            message=f"card '{label}' has {wb}-word body (max {ceiling} at 16pt for {layout})",
                            fix_hint=f"trim card '{label}' body to ≤{ceiling} words — drop the least important bullet or merge two related bullets",
                        ))
                # Title / label check
                title_field = card.get("title") or card.get("label") or ""
                if isinstance(title_field, str):
                    tw = _count_words(title_field)
                    if tw > title_ceiling:
                        issues.append(ValidationIssue(
                            rule="R11",
                            slide_index=i,
                            message=f"card title/label '{title_field}' has {tw} words (max {title_ceiling} at 20pt for {layout})",
                            fix_hint=f"shorten title/label to ≤{title_ceiling} words",
                        ))
    return issues


# =============================================================================
# PUBLIC ENTRY POINT
# =============================================================================

ALL_CHECKS = [
    _check_r1_deprecated_glyphs,
    _check_r2_paragraph_length,
    _check_r3_cards_2_large_prose,
    _check_r4_lavender_pair,
    _check_r5_lavender_density,
    _check_r6_filled_header_colors,
    _check_r7_icon_consistency,
    _check_r8_pair_diversity,
    _check_r9_pill_present,
    _check_r10_example_present,
    _check_r11_overflow,
]


def validate_plan(plan: dict) -> None:
    """Validate plan.json. Raises PlanValidationError if any hard rule fails.

    Fast-path plans (empty content_slides) skip validation entirely — the
    Empty template is delivered as-is so per-slide rules don't apply.
    """
    content_slides = plan.get("content_slides", [])
    if len(content_slides) == 0:
        return  # fast-path, no validation needed

    all_issues = []
    for check in ALL_CHECKS:
        all_issues.extend(check(content_slides))

    if all_issues:
        raise PlanValidationError(all_issues)


def validate_plan_collect(plan: dict) -> list:
    """Like validate_plan() but returns the issue list instead of raising.

    Used by auto-retry mode to inspect issues without exception-handling.
    """
    content_slides = plan.get("content_slides", [])
    if len(content_slides) == 0:
        return []
    issues = []
    for check in ALL_CHECKS:
        issues.extend(check(content_slides))
    return issues
