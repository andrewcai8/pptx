from deckedit import DeckEdit, variant

OLD_TITLE = "Significant impact was achieved"
NEW_TITLE = "Starbucks personalization lifted marketing engagement 150% and added $100M net incremental revenue in year 1"


@variant()
def good(d: DeckEdit) -> None:
    """Slide 15 carries an action title with the engagement and revenue figures; nothing else moves."""
    d.replace(15, OLD_TITLE, NEW_TITLE)


@variant(fails={("non-change", 14), ("missing", 15)})
def title_on_slide_14(d: DeckEdit) -> None:
    """The new action title lands on slide 14 instead of slide 15."""
    d.replace(14, "An illustrative view on the analytics engine", NEW_TITLE)


@variant(base=good, fails={("non-change", 14)})
def retitled_slide_14(d: DeckEdit) -> None:
    """Slide 14's title is rewritten this week although the client deferred it."""
    d.replace(14, "An illustrative view on the analytics engine", "Propensity scores rank which products to offer each customer")


@variant(fails={("forbidden", 15)})
def kept_tripled(d: DeckEdit) -> None:
    """The title keeps the client's first word, tripled, after she dropped it."""
    d.replace(15, OLD_TITLE, "Starbucks personalization tripled engagement: 150% more marketing engagement and $100M net incremental revenue in year 1")


@variant(fails={("style", 15)})
def title_too_long(d: DeckEdit) -> None:
    """The action title runs past the 150-character house limit."""
    d.replace(
        15,
        OLD_TITLE,
        "Starbucks one-to-one personalization lifted marketing engagement by 150% and delivered a 300% increase in net incremental revenue, worth $100M in year 1 of the program",
    )


@variant(base=good, fails={("scope", 15)})
def trimmed_quote(d: DeckEdit) -> None:
    """The first Starbucks quote is shortened although the client wanted the quotes verbatim."""
    d.replace(15, "within our app... Our digital flywheel momentum accelerated ... with the launch of", "within our app...")


@variant(base=good, fails={("non-change", 6)})
def cleaned_slide_6_dots(d: DeckEdit) -> None:
    """The deliberate four-dot continuation bullets on slide 6 are rewritten."""
    d.replace(6, "....", "Further use cases", count=4)


@variant()
def spelled_out_million(d: DeckEdit) -> None:
    """The action title writes the revenue as $100 million in year one, the same figure as the slide's +$100M."""
    d.replace(15, OLD_TITLE, "Starbucks personalization lifted marketing engagement by 150% and added $100 million in year one")
