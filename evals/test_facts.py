import pytest

from facts import match, parse


@pytest.mark.parametrize(
    ("kind", "spec", "text", "found"),
    [
        ("money", "$100M", "lifted engagement by 150% and added $100 million in year one", "$100 million"),
        ("money", "$100M", "added $100m in year 1", "$100m"),
        ("money", "$100M", "added US$100M in year 1", "us$100m"),
        ("money", "$100M", "added USD 100 million", "usd 100 million"),
        ("money", "$100M", "= +$100M in year 1", "$100m"),
        ("money", "$100M", "added $0.1bn in year 1", "$0.1bn"),
        ("money", "$410m", "The market is currently worth c.$410m in Africa", "$410m"),
        ("money", "$2bn", "grows to $2,000m by 2027", "$2,000m"),
        ("money", "$100M", "added $1,100M in year 1", None),
        ("money", "$100M", "added $100k in year 1", None),
        ("money", "$100M", "added €100M in year 1", None),
        ("money", "$100M", "added 100 million in year 1", None),
        ("money", "$0.07", "tariff of $0.07-0.11 per kWh", "$0.07"),
        ("money", "$100M", "added $100MM in year 1", "$100mm"),
        ("money", "$100M", "added 100 million dollars in year one", "100 million dollars"),
        ("money", "$100M", "Incremental revenue: 100m$", "100m$"),
        ("money", "$100M", "added US$ 100 million", "us$ 100 million"),
        ("money", "$100M", "added 100 million euros", None),
        ("money", "$100M", "our top 5 $100M deals", "$100m"),
        ("percent", "48%", "Millennials will account for ca. 48 per cent of the client base", "48 per cent"),
        ("percent", "48%", "ca. 48 percent of clients", "48 percent"),
        ("percent", "48%", "ca.48% of clients", "48%"),
        ("percent", "39%", "Potential +39%", "+39%"),
        ("percent", "39%", "a 139% rise", None),
        ("percent", "39%", "a 1.39% rise", None),
        ("percent", "39%", "a 39.5% rise", None),
        ("percent", "39%", "39 of the decks", None),
        ("percent", "48%", "a fall of -48% since 2017", None),
        ("percent", "48%", "a fall of −48% since 2017", None),
        ("percent", "48%", "from 50% to -48%", None),
        ("percent", "45%", "ca. 45-48% of the client base", "45-48%"),
        ("percent", "45%", "ca. 45–48% of the client base", "45–48%"),
        ("percent", "45%", "ca. 45% – 48% of the client base", "45% – 48%"),
        ("percent", "48%", "ca. 45-48% of the client base", "45-48%"),
        ("percent", "27%", "a CAGR over 2022-27 of 37%", None),
        ("count", "6 weeks", "STEP 1: a six-week diagnostic", "six-week"),
        ("count", "6 weeks", "Diagnose in 6 weeks", "6 weeks"),
        ("count", "6 weeks", "Diagnose in 6 calendar weeks", "6 calendar weeks"),
        ("count", "6 weeks", "Diagnose in 16 weeks", None),
        ("count", "6 weeks", "6 days, then weeks of planning", None),
        ("count", "10 levers", "All 10 value-creation levers assessed", "10 value-creation levers"),
        ("count", "10 levers", "8 of 10 levers", "10 levers"),
        ("count", "8 weeks", "8 of 10 levers in 6 weeks", None),
        ("count", "10 levers", "10 key value levers", "10 key value levers"),
        ("count", "10 levers", "all 10 of the levers", "10 of the levers"),
        ("count", "10 levers", "10 things we learned about the levers", None),
        ("count", "8 weeks", "8 days and two weeks", None),
        ("count", "6 weeks", "6 days and then weeks", None),
        ("count", "10 levers", "10 business process levers", "10 business process levers"),
        ("text", ["cap", "limit"], "Teams choose their days in the office, within the site's daily limit", "limit"),
        ("text", ["cap", "limit"], "within each site's capacity", None),
        ("text", "days", "Oversee day-to-day work", None),
        ("text", "Flexibility to teams", "FLEXIBILITY  TO TEAMS", "flexibility to teams"),
        ("text", ["team:", "# consultants"], "Team: 2 consultants", "team:"),
        ("text", "# consultants", "staffed with two consultants and a principal", "two consultants"),
        ("text", "# consultants", "a joint team of category leads and consultants", None),
        ("text", "team of #", "a joint team of Nordvik category leads", None),
        ("text", "team of #", "a team of 4", "team of 4"),
        ("text", "#-person", "a 12-person team", "12-person"),
        ("text", "# people", "1,200 people", None),
        ("text", "39%", "139% and +39%", "39%"),
    ],
)
def test_match_finds_a_fact_by_meaning(kind, spec, text, found):
    assert match(parse(kind, spec), text) == found


@pytest.mark.parametrize(
    ("kind", "spec"),
    [("money", "100 million"), ("percent", "45-48%"), ("money", "$100M and $5M"), ("percent", "48"), ("count", "weeks"), ("chart", "four hundred"), ("text", []), ("text", 5)],
)
def test_parse_rejects_a_spec_that_is_not_one_fact_of_its_kind(kind, spec):
    with pytest.raises(ValueError):
        parse(kind, spec)


def test_parse_compares_money_across_scales_and_currency_spellings():
    assert parse("money", "$0.4bn") == parse("money", "US$400 million") == parse("money", "$400m")
    assert parse("money", "$400m") != parse("money", "€400m")
