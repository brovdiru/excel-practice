"""Builds the static site from src/app.html. Run: python3 tools/build.py"""
import re, pathlib, datetime
ROOT = pathlib.Path(__file__).resolve().parent.parent
BASE = "https://brovdiru.github.io/excel-practice/"
SRC = (ROOT / "src/app.html").read_text()
TODAY = datetime.date.today().isoformat()

PAGES = [
  dict(file="index.html", topic=None, title="Excel Formula Practice Test",
       h1="Excel Formula Practice",
       desc=None, article=None),
  dict(file="vlookup-practice.html", topic=["lookup"], title="VLOOKUP and XLOOKUP Practice",
       h1="VLOOKUP, INDEX/MATCH and XLOOKUP practice",
       desc="Practice VLOOKUP, INDEX/MATCH, XLOOKUP and IFERROR on a real sales table. Type the formula, get instant feedback. Free, no sign-up.",
       article="""<h2>Lookup questions in Excel tests</h2>
<p>Almost every Excel assessment has at least one lookup task: find a price, a category or a stock level for a product code. Interviewers use it to see whether you know an exact match from an approximate one.</p>
<h3>VLOOKUP</h3>
<p><code>=VLOOKUP(D2,J2:L6,2,FALSE)</code> searches for D2 in the first column of J2:L6 and returns the value from the 2nd column. The last argument <code>FALSE</code> means exact match. Leaving it out is the most common mistake: Excel then does an approximate match and can return a wrong row without any error.</p>
<h3>INDEX and MATCH</h3>
<p><code>=INDEX(L2:L6,MATCH(D4,J2:J6,0))</code> does the same job and also works when the column you need is to the left of the lookup column. MATCH finds the position, INDEX returns the value at that position.</p>
<h3>XLOOKUP</h3>
<p>In Excel 365 and 2021, <code>=XLOOKUP(D3,J2:J6,L2:L6)</code> is the shortest version. It is exact by default and has a built-in fourth argument for a "not found" value.</p>
<h3>Handling #N/A</h3>
<p>When the value is missing, lookups return <code>#N/A</code>. Wrap the formula in <code>IFERROR(…,"Not found")</code> to show a clean message.</p>"""),
  dict(file="sumif-countif-practice.html", topic=["cond"], title="SUMIF and COUNTIF Practice",
       h1="SUMIF, COUNTIF and SUMIFS practice",
       desc="Interactive SUMIF, COUNTIF, AVERAGEIF, SUMIFS, COUNTIFS and MAXIFS exercises with instant checking. Prepare for Excel interview tests.",
       article="""<h2>Conditional sums and counts</h2>
<p>"Total sales for the West region" or "how many orders were over 1,000" are standard Excel test questions. They check that you can filter inside a formula instead of by hand.</p>
<h3>COUNTIF and SUMIF</h3>
<p><code>=COUNTIF(B2:B13,"East")</code> counts matching rows. <code>=SUMIF(B2:B13,"West",G2:G13)</code> checks column B and adds up column G. Conditions with operators go in quotes: <code>"&gt;1000"</code>, <code>"&lt;&gt;East"</code>.</p>
<h3>SUMIFS and COUNTIFS</h3>
<p>For several conditions use the plural versions. Note the order: in SUMIFS the sum range comes <em>first</em>, <code>=SUMIFS(G2:G13,B2:B13,"East",D2:D13,"Laptop")</code>. Mixing up SUMIF and SUMIFS argument order is a frequent interview slip.</p>
<h3>AVERAGEIF, MAXIFS, MINIFS</h3>
<p>Same pattern, different result: average, largest or smallest value among matching rows.</p>
<h3>Wildcards</h3>
<p>Text criteria accept <code>*</code> (any characters) and <code>?</code> (one character), so <code>"Lap*"</code> matches Laptop.</p>"""),
  dict(file="if-function-practice.html", topic=["logic"], title="Excel IF Function Practice",
       h1="Excel IF, AND and OR practice",
       desc="Practice the Excel IF function with AND, OR and percentages on a real dataset. Instant feedback on every formula. Free.",
       article="""<h2>IF formulas in Excel tests</h2>
<p>IF questions test logic more than syntax: label orders as "Big" or "Small", flag rows for review, or calculate a commission only when a threshold is met.</p>
<h3>The basic shape</h3>
<p><code>=IF(G2&gt;=2000,"Big","Small")</code>: a test, the value if true, the value if false. Text results need double quotes. Numbers do not.</p>
<h3>Several conditions</h3>
<p>Put <code>AND(…)</code> or <code>OR(…)</code> inside the test: <code>=IF(OR(E10&gt;12,G10&lt;400),"Check","")</code>. AND needs every condition to be true, OR needs at least one.</p>
<h3>Calculations inside IF</h3>
<p><code>=IF(G7&gt;2000,G7*5%,0)</code> returns a commission or zero. Excel understands <code>5%</code> as 0.05.</p>
<h3>Common mistakes</h3>
<ul><li>Forgetting quotes around text, which gives <code>#NAME?</code>.</li><li>Using <code>&gt;</code> when the task says "or more" (that needs <code>&gt;=</code>).</li><li>Typing the result by hand instead of referring to the cell.</li></ul>"""),
  dict(file="excel-text-functions-practice.html", topic=["text"], title="Excel Text Functions Practice",
       h1="Excel text functions practice",
       desc="Practice LEFT, RIGHT, MID, LEN, FIND, UPPER and joining text with & in Excel. Split names and build codes with instant feedback.",
       article="""<h2>Text tasks in Excel tests</h2>
<p>Splitting a full name into first and last name, or building a code from several columns, shows up in admin, HR and data-entry assessments.</p>
<h3>First name</h3>
<p><code>=LEFT(C2,FIND(" ",C2)-1)</code>. FIND returns the position of the space, and LEFT takes everything before it. In Excel 365 <code>=TEXTBEFORE(C2," ")</code> does the same.</p>
<h3>Last name</h3>
<p><code>=RIGHT(C6,LEN(C6)-FIND(" ",C6))</code>. The length of the name minus the position of the space gives the number of characters to take from the right.</p>
<h3>Joining text</h3>
<p>The <code>&amp;</code> operator joins values: <code>=UPPER(B2)&amp;"-"&amp;A2</code> turns East and 1001 into EAST-1001. CONCAT and TEXTJOIN work too.</p>"""),
  dict(file="excel-basic-formulas-practice.html", topic=["basics","adv"], title="Basic Excel Formulas Practice",
       h1="Basic Excel formulas practice",
       desc="Practice SUM, AVERAGE, MIN, MAX, COUNT, ROUND, SUMPRODUCT and absolute references ($) in a real spreadsheet. Free interactive exercises.",
       article="""<h2>The basics every Excel test starts with</h2>
<p>The first questions of most assessments are quick points: total, average, largest and smallest values, count of rows. Get them fast and right to save time for harder tasks.</p>
<h3>SUM, AVERAGE, MIN, MAX, COUNT</h3>
<p>All take a range: <code>=SUM(G2:G13)</code>. Check that the range covers every row of data, not one less.</p>
<h3>ROUND</h3>
<p><code>=ROUND(AVERAGE(G2:G13),0)</code> rounds to whole numbers. Use 2 for cents.</p>
<h3>Absolute references</h3>
<p>The <code>$</code> sign fixes a reference when you copy a formula. <code>=G2/SUM($G$2:$G$13)</code> keeps the total range in place in every row, while G2 moves to G3, G4 and so on.</p>
<h3>SUMPRODUCT</h3>
<p><code>=SUMPRODUCT(E2:E13,F2:F13)</code> multiplies units by price row by row and adds the results, all in one cell.</p>"""),
]

def nav(cur):
    items = "".join(f'<li><a href="{p["file"]}">{p["title"]}</a></li>' for p in PAGES if p["file"] not in (cur, "index.html"))
    return f'<h2>More practice</h2><ul><li><a href="./">All 29 exercises and the mock interview test</a></li>{items}</ul>'

HEAD = '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
for p in PAGES:
    html = SRC
    url = BASE + ("" if p["file"] == "index.html" else p["file"])
    html = re.sub(r"<title>.*?</title>", f"<title>{p['title']}</title>", html, count=1)
    if p["desc"]:
        html = re.sub(r'<meta name="description" content="[^"]*">', f'<meta name="description" content="{p["desc"]}">', html, count=1)
        html = re.sub(r"<h1>.*?</h1>", f"<h1>{p['h1']}</h1>", html, count=1)
        html = re.sub(r'<article class="article">.*?</article>', f'<article class="article">{p["article"]}{nav(p["file"])}</article>', html, count=1, flags=re.S)
        html = re.sub(r'<script type="application/ld\+json">.*?</script>\n', "", html, count=1, flags=re.S)
        html = html.replace("<script>\n/* ================= Formula engine", f"<script>window.PAGE_TOPIC={p['topic']!r};</script>\n<script>\n/* ================= Formula engine".replace("'", '"'), 1)
    else:
        html = html.replace("</article>", nav("index.html") + "</article>", 1)
    html = html.replace("<footer>", '<footer><a href="about.html">About</a> · <a href="privacy.html">Privacy</a> · ', 1)
    (ROOT / p["file"]).write_text(HEAD + f'<link rel="canonical" href="{url}">\n' + html + "</html>\n")

urls = [BASE + ("" if p["file"] == "index.html" else p["file"]) for p in PAGES] + [BASE + "about.html", BASE + "privacy.html"]
(ROOT / "sitemap.xml").write_text('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
    + "".join(f"  <url><loc>{u}</loc><lastmod>{TODAY}</lastmod></url>\n" for u in urls) + "</urlset>\n")
print("built", len(PAGES), "pages")
