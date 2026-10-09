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


@variant(base=good, fails={("lost", 15)})
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


@variant()
def revenue_in_dollars_after(d: DeckEdit) -> None:
    """The action title writes the revenue as 100 million dollars, the same figure as the slide's +$100M."""
    d.replace(15, OLD_TITLE, "Starbucks personalization lifted marketing engagement by 150% and added 100 million dollars in year one")


@variant()
def hyphenated_100_million(d: DeckEdit) -> None:
    """The action title writes the revenue as a $100-million uplift, the same figure as the slide's +$100M."""
    d.replace(15, OLD_TITLE, "Starbucks personalization lifted marketing engagement by 150%, a $100-million uplift in year one")


@variant()
def title_split_into_two_paragraphs(d: DeckEdit) -> None:
    """The action title is split into two paragraphs, as pressing Enter in PowerPoint gives; Pieter allows two lines."""
    d.slide(15).shapes.title.text_frame.text = "Starbucks personalisation lifted marketing engagement by 150%\nand added $100M net revenue in year one"


@variant(base=good, fails={("scope", 5), ("forbidden", 5)})
def tripled_on_another_slide(d: DeckEdit) -> None:
    """The title is right, but slide 5's 3x revenue callout now says "Tripled", a word Ines asked to see nowhere."""
    d.replace(5, "3x", "Tripled")


@variant(base=good, fails={("forbidden", 15)})
def tripled_in_speaker_notes(d: DeckEdit) -> None:
    """The title is right, but slide 15's speaker notes say engagement tripled, and the board gets the pptx with its notes."""
    d.slide(15).notes_slide.notes_text_frame.text = "Talk track: personalization tripled engagement in the Starbucks app."


@variant(base=good)
def talk_track_in_speaker_notes(d: DeckEdit) -> None:
    """Slide 15 gains speaker notes that walk through the 150% and the $100M without the word Ines ruled out."""
    d.slide(15).notes_slide.notes_text_frame.text = "Talk track: lead with the 150% engagement lift, then the $100M net incremental revenue in year 1."
