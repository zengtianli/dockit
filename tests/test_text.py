"""Tests for dockit.text module."""

from dockit.text import fix_all, fix_punctuation, fix_quotes, fix_units


class TestFixQuotes:
    def test_basic_pair(self):
        result, count, counter = fix_quotes('"hello"')
        assert result == "\u201chello\u201d"
        assert count == 2
        assert counter == 2

    def test_counter_continuity(self):
        """Cross-run counter should maintain pairing."""
        r1, _, counter = fix_quotes('"hello', counter=0)
        assert r1 == "\u201chello"
        assert counter == 1

        r2, _, counter = fix_quotes('world"', counter=counter)
        assert r2 == "world\u201d"
        assert counter == 2

    def test_no_quotes(self):
        result, count, counter = fix_quotes("no quotes here")
        assert result == "no quotes here"
        assert count == 0

    def test_mixed_quote_types(self):
        """Should handle various quote characters."""
        result, count, _ = fix_quotes('\u201chello\u201d')
        assert count == 2
        assert "\u201c" in result
        assert "\u201d" in result


class TestFixPunctuation:
    def test_comma(self):
        result, count = fix_punctuation("hello, world")
        assert result == "hello\uff0c world"
        assert count == 1

    def test_multiple(self):
        result, count = fix_punctuation("what? yes! ok:")
        assert result == "what\uff1f yes\uff01 ok\uff1a"
        assert count == 3

    def test_parentheses(self):
        result, count = fix_punctuation("(test)")
        assert result == "\uff08test\uff09"
        assert count == 2

    def test_no_change(self):
        result, count = fix_punctuation("\u5df2\u7ecf\u662f\u4e2d\u6587\u6807\u70b9\uff0c")
        assert count == 0


class TestFixUnits:
    def test_area(self):
        result, count = fix_units("\u5e73\u65b9\u7c73")
        assert result == "m\u00b2"
        assert count == 1

    def test_volume(self):
        result, count = fix_units("\u7acb\u65b9\u7c73")
        assert result == "m\u00b3"
        assert count == 1

    def test_temperature(self):
        result, count = fix_units("25\u6444\u6c0f\u5ea6")
        assert result == "25\u2103"
        assert count == 1

    def test_longer_match_first(self):
        """Should match \u5e73\u65b9\u516c\u91cc before \u516c\u91cc."""
        result, count = fix_units("100\u5e73\u65b9\u516c\u91cc")
        assert result == "100km\u00b2"
        assert count == 1

    def test_superscript(self):
        result, count = fix_units("100m2")
        assert result == "100m\u00b2"
        assert count == 1


class TestFixAll:
    def test_combined(self):
        text = '"\u6d4b\u8bd5\u6587\u672c",\u9762\u79ef100\u5e73\u65b9\u7c73'
        result, stats, counter = fix_all(text)
        assert stats["quotes"] == 2
        assert stats["punctuation"] == 1
        assert stats["units"] == 1
        assert "m\u00b2" in result
        assert "\u201c" in result
        assert "\uff0c" in result

    def test_empty_string(self):
        result, stats, counter = fix_all("")
        assert result == ""
        assert stats == {"quotes": 0, "punctuation": 0, "units": 0}
