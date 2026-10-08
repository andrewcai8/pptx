from pptx.util import Emu, Pt

from deckedit import DeckEdit, variant

TITLE = "STEP 1: Diagnose in 6 weeks, ready for the 2 December steering committee"
BULLETS = (
    "6 weeks from kick-off on 19 October to readout",
    "All 10 levers assessed against peers, across both pages of the imperatives house",
    "Deliverables: peer benchmark and a heatmap of where you are ahead, behind or at risk on each lever",
    "Readout at the steering committee on 2 December",
)


NO_PLACEHOLDER = ("n4", "no team-size placeholder")


# The layout's level 0 renders as unbulleted 10pt text, unlike the 16 to 18pt bullets on the deck's own text slides.
def add_diagnostic(d: DeckEdit, *, after: int = 13, title: str = TITLE, bullets=BULLETS) -> None:
    d.add_slide("c1", after=after, layout="D. Title and Text", title=title, bullets=bullets)
    body = next(ph for ph in d.slide("c1").placeholders if ph.placeholder_format.idx != 0)
    for p in body.text_frame.paragraphs:
        p.level = 1
        for r in p.runs:
            r.font.size = Pt(16)


@variant()
def good(d: DeckEdit) -> None:
    """Diagnostic slide on the deck's own text layout, straight after Next steps, with 6 weeks and 10 levers."""
    add_diagnostic(d)


@variant(fails={("structure", "c1")})
def appended_at_end(d: DeckEdit) -> None:
    """Diagnostic slide appended after the last slide instead of straight after Next steps."""
    add_diagnostic(d, after=16)


@variant(fails={("structure", "c1")})
def placed_after_contacts(d: DeckEdit) -> None:
    """Diagnostic slide placed after the contacts page, the placement Nadia talked the client out of."""
    add_diagnostic(d, after=14)


@variant(fails={("layout", "c1")})
def title_only_with_textbox(d: DeckEdit) -> None:
    """Diagnostic slide built on D. Title Only with the bullets in a hand-drawn text box."""
    d.add_slide("c1", after=13, layout="D. Title Only", title=TITLE)
    slide = d.slide("c1")
    title = slide.shapes.title
    box = slide.shapes.add_textbox(title.left, Emu(title.top + title.height + Pt(18)), title.width, Pt(260))
    tf = box.text_frame
    tf.text = BULLETS[0]
    for b in BULLETS[1:]:
        tf.add_paragraph().text = b


@variant(fails={("style", "c1")})
def tbd_team_bullet(d: DeckEdit) -> None:
    """Diagnostic slide carries a "Team size TBD" placeholder bullet the client asked to leave off; the house style flags the TBD."""
    add_diagnostic(d, bullets=BULLETS + ("Team size TBD",))


@variant(fails={("missing", "c1"), ("forbidden", "c1")})
def kept_eight_weeks(d: DeckEdit) -> None:
    """Diagnostic slide keeps the superseded 8 weeks instead of the 6 weeks the client settled on."""
    add_diagnostic(d, title=TITLE.replace("6 weeks", "8 weeks"), bullets=tuple(b.replace("6 weeks", "8 weeks") for b in BULLETS))


@variant(base=good, fails={("non-change", 14)})
def edited_contacts(d: DeckEdit) -> None:
    """Contacts page edited to swap an original author for the consultant, which Nadia parked until marketing has been asked."""
    d.replace(14, "Matt Gamber", "Nadia Brennan")


@variant(fails={("structure", "c1"), ("non-change", 14)})
def after_contacts_and_edited_contacts(d: DeckEdit) -> None:
    """Diagnostic slide placed after the contacts page, and the contacts page edited to name the consultant."""
    add_diagnostic(d, after=14)
    d.replace(14, "Matt Gamber", "Nadia Brennan")


@variant()
def joint_team_named(d: DeckEdit) -> None:
    """Diagnostic slide says it is run by a joint team of Nordvik category leads and consultants, which names who works on it but no size."""
    add_diagnostic(d, bullets=BULLETS + ("Run by a joint team of Nordvik category leads and consultants",))


@variant(fails={("non-change", "c1")})
def team_of_two_consultants(d: DeckEdit) -> None:
    """Diagnostic slide carries "Team: 2 consultants", a team size the client said to leave off."""
    add_diagnostic(d, bullets=BULLETS + ("Team: 2 consultants",))


@variant(intent=NO_PLACEHOLDER)
def team_tbc_with_lars(d: DeckEdit) -> None:
    """Diagnostic slide carries "Team: TBC with Lars", the placeholder the client said not to put."""
    add_diagnostic(d, bullets=BULLETS + ("Team: TBC with Lars",))


@variant(fails={("non-change", "c1")})
def staffed_two_consultants(d: DeckEdit) -> None:
    """Diagnostic slide says it is staffed with two consultants and a principal, a team size in words."""
    add_diagnostic(d, bullets=BULLETS + ("Staffed with two consultants and a principal",))


@variant()
def levers_with_words_between(d: DeckEdit) -> None:
    """Diagnostic slide says "all 10 of the value levers", the same count with words between the number and the noun."""
    add_diagnostic(d, bullets=(BULLETS[0], "All 10 of the value levers assessed against peers", *BULLETS[2:]))


@variant(fails={("non-change", "c1")})
def team_of_two_bcg_consultants(d: DeckEdit) -> None:
    """Diagnostic slide carries "Team: 2 BCG consultants", a team size with a word between the number and the noun."""
    add_diagnostic(d, bullets=BULLETS + ("Team: 2 BCG consultants",))


@variant(fails={("non-change", "c1")})
def two_senior_consultants(d: DeckEdit) -> None:
    """Diagnostic slide says it is run by 2 senior consultants, a team size the client said to leave off."""
    add_diagnostic(d, bullets=BULLETS + ("Run by 2 senior consultants",))


@variant(intent=NO_PLACEHOLDER)
def resourcing_to_be_confirmed(d: DeckEdit) -> None:
    """Diagnostic slide carries "Resourcing: to be confirmed", the placeholder the client said not to put."""
    add_diagnostic(d, bullets=BULLETS + ("Resourcing: to be confirmed",))


@variant(intent=NO_PLACEHOLDER)
def staffing_tbc_with_lars(d: DeckEdit) -> None:
    """Diagnostic slide carries "Staffing: TBC with Lars", a team placeholder in another wording."""
    add_diagnostic(d, bullets=BULLETS + ("Staffing: TBC with Lars",))


@variant()
def staffed_jointly(d: DeckEdit) -> None:
    """Diagnostic slide says it is staffed jointly by Nordvik category leads and consultants, which names who but no size."""
    add_diagnostic(d, bullets=BULLETS + ("Staffed jointly by Nordvik category leads and consultants",))


@variant()
def kickoff_tbc(d: DeckEdit) -> None:
    """Diagnostic slide marks the 19 October kick-off as TBC, a placeholder that is not about the team."""
    add_diagnostic(d, bullets=("6 weeks from kick-off on 19 October (TBC) to readout", *BULLETS[1:]))


@variant(fails={("non-change", "c1")})
def counted_promo_planners(d: DeckEdit) -> None:
    """Diagnostic slide says it draws on the 12 people who run Q4 promo planning, a headcount the client never gave and kept for the budget talk."""
    add_diagnostic(d, bullets=BULLETS + ("Draws on the 12 people who run Q4 promo planning",))


@variant()
def readout_by_the_consultants(d: DeckEdit) -> None:
    """Diagnostic slide says the readout on 2 December is given by the consultants; a date before the noun is not a count."""
    add_diagnostic(d, bullets=(*BULLETS[:3], "Readout on 2 December by the consultants at the steering committee"))


@variant()
def weeks_1_to_6(d: DeckEdit) -> None:
    """Diagnostic slide states the duration only as "Weeks 1 to 6", which runs the 6 weeks from week one."""
    add_diagnostic(d, title="STEP 1: Diagnose, ready for the 2 December steering committee", bullets=("Weeks 1 to 6: from kick-off on 19 October to readout", *BULLETS[1:]))


@variant(intent=NO_PLACEHOLDER)
def placeholder_for_team(d: DeckEdit) -> None:
    """Diagnostic slide carries "Team (TBC)", a placeholder for the team size the client said to leave off."""
    add_diagnostic(d, bullets=BULLETS + ("Team (TBC)",))


@variant(fails={("forbidden", "c1")})
def hedged_6_to_8_weeks(d: DeckEdit) -> None:
    """Diagnostic slide hedges the duration as "6 to 8 weeks", keeping the 8 weeks the client ruled out."""
    add_diagnostic(d, bullets=("6 to 8 weeks from kick-off on 19 October to readout", *BULLETS[1:]))


@variant()
def step_label_with_consultants(d: DeckEdit) -> None:
    """Diagnostic slide says "Step 1 run by consultants alongside Nordvik category leads"; the 1 labels the step, it counts nobody."""
    add_diagnostic(d, bullets=BULLETS + ("Step 1 run by consultants alongside Nordvik category leads",))


@variant()
def phase_label_with_consultants(d: DeckEdit) -> None:
    """Diagnostic slide says "Phase 1 led by the consultants", a phase label and no team size."""
    add_diagnostic(d, bullets=BULLETS + ("Phase 1 led by the consultants",))


@variant()
def day_label_with_consultants(d: DeckEdit) -> None:
    """Diagnostic slide says "Day 1 workshop with the consultants", a day label and no team size."""
    add_diagnostic(d, bullets=BULLETS + ("Day 1 workshop with the consultants",))


@variant()
def step_title_with_consultants(d: DeckEdit) -> None:
    """Diagnostic slide is titled "Step 1 Diagnose with your consultants: 6 weeks, all 10 levers", which names the step and no team size."""
    add_diagnostic(d, title="Step 1 Diagnose with your consultants: 6 weeks, all 10 levers")


@variant()
def duration_in_wks(d: DeckEdit) -> None:
    """Diagnostic slide states the duration only as "Duration 6 wks", the usual short form of 6 weeks."""
    add_diagnostic(d, title="STEP 1: Diagnose, ready for the 2 December steering committee", bullets=("Duration 6 wks, from kick-off on 19 October to readout", *BULLETS[1:]))


@variant()
def levers_covered_label(d: DeckEdit) -> None:
    """Diagnostic slide states the lever count only as the label "Levers covered: all 10"."""
    add_diagnostic(d, bullets=(BULLETS[0], "Levers covered: all 10", *BULLETS[2:]))


@variant(fails={("forbidden", "c1")})
def extended_option_8_wks(d: DeckEdit) -> None:
    """Diagnostic slide keeps "Extended option: 8 wks", the eight Nadia said to take out completely."""
    add_diagnostic(d, bullets=BULLETS + ("Extended option: 8 wks",))


@variant(intent=NO_PLACEHOLDER)
def number_of_consultants_tbc(d: DeckEdit) -> None:
    """Diagnostic slide carries "Number of consultants: TBC", a team placeholder with no number for a script to read."""
    add_diagnostic(d, bullets=BULLETS + ("Number of consultants: TBC",))


@variant(intent=NO_PLACEHOLDER)
def people_tbc(d: DeckEdit) -> None:
    """Diagnostic slide carries "People: TBC", a team placeholder with no number for a script to read."""
    add_diagnostic(d, bullets=BULLETS + ("People: TBC",))


@variant(intent=NO_PLACEHOLDER)
def staffing_to_be_agreed_with_lars(d: DeckEdit) -> None:
    """Diagnostic slide carries "Staffing: to be agreed with Lars", a team placeholder with no number for a script to read."""
    add_diagnostic(d, bullets=BULLETS + ("Staffing: to be agreed with Lars",))


@variant(intent=NO_PLACEHOLDER)
def team_of_x_consultants(d: DeckEdit) -> None:
    """Diagnostic slide carries "Team: [x] consultants", a team placeholder with no number for a script to read."""
    add_diagnostic(d, bullets=BULLETS + ("Team: [x] consultants",))
