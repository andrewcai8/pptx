import copy

from pptx.enum.text import MSO_AUTO_SIZE

from deckedit import DeckEdit, variant

SURVEY_NOTE = "3. Insurers' Association member survey 2026, 2027 projection"
CENSUS_NOTE = "1. U.S. Census Bureau, Population Division.  2014 estimate of population; Generations as defined by Pew Research Center, 2014"


def footnote_paragraphs(d: DeckEdit) -> list:
    return [p._p for p in next(s for s in d.slide(5).shapes if s.name == "Footnote").text_frame.paragraphs]


def cite_survey(d: DeckEdit, *, text: str = SURVEY_NOTE, first: bool = False) -> None:
    notes = footnote_paragraphs(d)
    note = copy.deepcopy(notes[-1])
    run, *rest = note.r_lst
    run.text = text
    for r in rest:
        note.remove(r)
    if first:
        notes[0].addprevious(note)
    else:
        notes[-1].addnext(note)


def refresh(d: DeckEdit) -> None:
    d.replace(5, "ca. 50%", "ca. 48%")
    d.delete_slide(2)
    d.move_slide(14, after=11)


@variant()
def good(d: DeckEdit) -> None:
    """Survey figure in the headline with a new footnote line citing the survey, credentials page gone, innovation leads the three roles."""
    refresh(d)
    cite_survey(d)


@variant()
def survey_added_to_footnote_1(d: DeckEdit) -> None:
    """Footnote 1 is extended to cite the survey for the headline, the literal reading of "I'll point the footnote at your survey file"."""
    refresh(d)
    d.replace(5, CENSUS_NOTE, CENSUS_NOTE + "; headline share: Insurers' Association member survey (2027 projection)")


@variant()
def survey_footnote_listed_first(d: DeckEdit) -> None:
    """The survey citation is a new footnote line above the census one, the other reading of pointing the footnote at the survey file."""
    refresh(d)
    cite_survey(d, text="Headline share: Insurers' Association member survey 2026 (2027 projection)", first=True)


@variant(base=good, fails={("scope", 5)})
def fixed_callout_spelling(d: DeckEdit) -> None:
    """Beyond the headline and footnote, slide 5's callout gets "Millenials" corrected, a line no change asks about."""
    d.replace(5, "...and Millenials have", "...and Millennials have")


@variant(base=good, fails={("guessed", 18)})
def flag_in_speaker_note(d: DeckEdit) -> None:
    """The open question about the regulator slide is written as a speaker note on slide 18 instead of in flags.json, so it is an edit."""
    d.slide(18).notes_slide.notes_text_frame.text = "Marko: which regulator slide should be punchier? Check before editing."


@variant(base=good, fails={("guessed", 18)})
def guessed_regulator_slide(d: DeckEdit) -> None:
    """Picks slide 18 as "the regulator slide" and rewrites it instead of asking which one."""
    d.replace(18, "have to closely cooperate", "must co-own the digital agenda")


@variant(fails={("missing", 5), ("forbidden", 5)})
def kept_first_answer(d: DeckEdit) -> None:
    """Uses the 45% Marko remembered and ignores the survey file he read out later."""
    d.replace(5, "ca. 50%", "ca. 45%")
    d.delete_slide(2)
    d.move_slide(14, after=11)


@variant(fails={("forbidden", 5)})
def hedged_range_45_to_48(d: DeckEdit) -> None:
    """The headline hedges with "ca. 45-48%", keeping the 45% Marko took back once he had the survey file."""
    refresh(d)
    d.replace(5, "ca. 48%", "ca. 45-48%")
    cite_survey(d)


@variant(fails={("structure", 2)})
def kept_credentials(d: DeckEdit) -> None:
    """Applies the number and the move but leaves the credentials page in."""
    d.replace(5, "ca. 50%", "ca. 48%")
    d.move_slide(14, after=11)


@variant(base=good, fails={("scope", 1), ("non-change", 1)})
def acted_on_logistics(d: DeckEdit) -> None:
    """Puts the Oversight Office on the cover although the invitation is the client's action and unconfirmed."""
    d.replace(1, "Ljubljana, September 4th, 2017", "Joint workshop with the Oversight Office, Ljubljana")


@variant(base=good, fails={("non-change", 12), ("non-change", 13), ("non-change", 14)})
def renumbered_roles(d: DeckEdit) -> None:
    """Renumbers the role pages after the move although the client said to keep the numbers."""
    d.replace(14, "3", "1")
    d.replace(12, "1", "2")
    d.replace(13, "2", "3")


@variant(base=good, fails={("scope", 9)})
def restyled_frozen_slide(d: DeckEdit) -> None:
    """Enlarges the title on slide 9, which nobody asked about; the text is the same, the XML is not."""
    d.set_size(9, "There are typical challenges", 32)


@variant(base=good, fails={("style", 5)})
def shrunk_footnote(d: DeckEdit) -> None:
    """Shrinks the slide 5 survey footnote to 6pt, below the house minimum."""
    d.set_size(5, "BCG digital satisfaction survey", 6)


@variant(base=good, fails={("structure", 18), ("guessed", 18)})
def dropped_regulator_slide(d: DeckEdit) -> None:
    """Deletes the closing regulator slide as a way of making "the regulator slide" punchier."""
    d.delete_slide(18)


@variant(fails={("structure", 14)})
def moved_after_wrong_slide(d: DeckEdit) -> None:
    """Puts genuine innovation after Digitize the core instead of straight after the roles overview."""
    d.replace(5, "ca. 50%", "ca. 48%")
    d.delete_slide(2)
    d.move_slide(14, after=12)


@variant(base=good, fails={("structure", 5)})
def moved_edited_slide(d: DeckEdit) -> None:
    """The generations slide gets the survey figure and is also moved after slide 6, which nobody asked for."""
    d.move_slide(5, after=6)


@variant(base=good, fails={("guessed", 18)})
def cleared_regulator_panel(d: DeckEdit) -> None:
    """The white panel behind slide 18's cartoon is made transparent, a guess at making "the regulator slide" land harder."""
    next(s for s in d.slide(18).shapes if s.name == "Rectangle 5").fill.background()


@variant(base=good, fails={("non-change", 17)})
def autofit_sandbox_title(d: DeckEdit) -> None:
    """Slide 17's title is set to shrink on overflow, which renders it smaller on the sandbox page Marko called fine."""
    d.slide(17).shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
