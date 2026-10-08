from deckedit import DeckEdit, variant

OLD = "Flexibility to teams"
NEW = "Teams choose their own office days within the site cap"


@variant()
def good(d: DeckEdit) -> None:
    """Slide 4 flexibility bullet reworded to the wording Daniel settled on."""
    d.replace(4, OLD, NEW)


@variant(fails={("scope", 6), ("missing", 4)})
def wrong_slide(d: DeckEdit) -> None:
    """The new wording overwrites the flexible working model bullet on slide 6 and slide 4 keeps the old bullet."""
    d.replace(
        6,
        "Working model will also need to be flexible during re-opening, with ability to moderate up/down, based on local risk factors (e.g., virus resurgence)",
        NEW,
    )


@variant(base=good, fails={("style", 4)})
def trailing_period(d: DeckEdit) -> None:
    """The reworded bullet ends with a full stop, a new no-bullet-end-punctuation breach."""
    d.replace(4, NEW, NEW + ".")


@variant(base=good, fails={("non-change", 7)})
def edited_timing_slide(d: DeckEdit) -> None:
    """Slide 7 gets a provincial-announcement caveat after the team agreed to leave it."""
    d.replace(
        7,
        "Government regulations and company guiding principles should determine timing for re-opening offices",
        "Government regulations and company guiding principles should determine timing for re-opening offices (pending provincial announcement)",
    )


@variant(fails={("missing", 4), ("forbidden", 4)})
def kept_first_answer(d: DeckEdit) -> None:
    """Slide 4 uses Daniel's first wording, which the team talked him out of."""
    d.replace(4, OLD, "Let teams pick their own schedule")


@variant(base=good, fails={("non-change", 1)})
def stamped_cover(d: DeckEdit) -> None:
    """The cover is stamped as a draft though Daniel said to leave it alone."""
    d.replace(1, "Edition #1: June 2020", "Edition #1: June 2020 | Draft for ExCo")


@variant(base=good, fails={("scope", 4)})
def dropped_profitability_bullet(d: DeckEdit) -> None:
    """The flexibility bullet is reworded, and the business profitability bullet next to it is emptied."""
    d.replace(4, "Business profitability, market share, etc.", "")
