#!/usr/bin/env python3
"""Check Treasury entries against the taxonomy and rebuild topic-page entry lists.

Usage (from the treasury repo root):
  python3 _tools/sync.py           rebuild topic pages' notes and the homepage
                                   recent list from entry dates, then report
  python3 _tools/sync.py --check   report only; exit 1 if anything is broken

Stdlib only. The taxonomy is the `topic_groups:` list in index.md.
"""
import datetime
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
MONTHS = {m: i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"], 1)}


def slugify(t):
    return re.sub(r"[^a-z0-9]+", "-", t.lower()).strip("-")


def unquote(v):
    v = v.strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        v = v[1:-1].replace('\\"', '"').replace("\\\\", "\\")
    return v


def q(s):
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def split_fm(text):
    m = re.match(r"---\n(.*?)\n---\n", text, re.S)
    if not m:
        return None, text
    return m.group(1), text[m.end():]


def parse_date(v):
    v = unquote(str(v))
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})$", v)
    if m:
        return datetime.date(*map(int, m.groups()))
    m = re.match(r"(\d{1,2}) ([A-Z][a-z]+) (\d{4})$", v)
    if m and m.group(2) in MONTHS:
        return datetime.date(int(m.group(3)), MONTHS[m.group(2)], int(m.group(1)))
    return None


def short(d):
    return f"{d.day} {d.strftime('%b')}"


def load_entries(errors):
    entries = {}
    for f in sorted(glob.glob("_pages/*.md")):
        slug = os.path.basename(f)[:-3]
        fm, _ = split_fm(open(f, encoding="utf-8").read())
        if fm is None:
            errors.append(f"{f}: no front matter")
            continue
        e = {"slug": slug, "file": f, "related": []}
        for key in ("id", "title", "updated", "created"):
            m = re.search(rf"^{key}:\s*(.*)$", fm, re.M)
            e[key] = unquote(m.group(1)) if m else None
        m = re.search(r"^topics:\s*\[(.*)\]\s*$", fm, re.M)
        e["topics"] = [unquote(t) for t in m.group(1).split(",") if t.strip()] if m else []
        m = re.search(r"^related:\n((?:[ \t-].*\n?)*)", fm, re.M)
        if m:
            e["related"] = re.findall(r'url:\s*"?/?([^"/\n]+)/?"?', m.group(1))
        e["date"] = parse_date(e["updated"] or "")
        e["cdate"] = parse_date(e["created"] or "") or e["date"]
        entries[slug] = e
    return entries


def load_taxonomy():
    fm, _ = split_fm(open("index.md", encoding="utf-8").read())
    groups, cur = [], None
    block = re.search(r"^topic_groups:\n((?:(?:[ \t].*)?\n)*)", fm + "\n", re.M).group(1)
    for line in block.splitlines():
        m = re.match(r"\s*- name:\s*(.*)$", line)
        if m:
            cur = (unquote(m.group(1)), [])
            groups.append(cur)
            continue
        m = re.match(r"\s{6,}- (.*)$", line)
        if m and cur:
            cur[1].append(unquote(m.group(1)))
    recent = re.findall(r'^  - title: .*\n    url: "?/?([^"/\n]+)/?"?\n    date: "?([^"\n]+)"?',
                        re.search(r"^recent:\n((?:(?:[ \t].*)?\n)*)", fm + "\n", re.M).group(1), re.M)
    m = re.search(r"^recent_exclude:\n((?:[ \t]+- .*\n?)*)", fm, re.M)
    exclude = [unquote(x) for x in re.findall(r"- (.*)", m.group(1))] if m else []
    exclude = [x.strip("/") for x in exclude]
    return groups, recent, exclude


def rebuild_recent(entries, recent, exclude):
    """Rewrite index.md's recent: list from entries' updated dates, newest first."""
    prev = {u: i for i, (u, _) in enumerate(recent)}
    order = sorted((s for s in entries if s not in exclude),
                   key=lambda s: (-entries[s]["date"].toordinal(), prev.get(s, -1), entries[s]["title"]))
    block = "recent:\n" + "".join(
        f'  - title: {q(entries[s]["title"])}\n    url: "/{s}/"\n    date: "{short(entries[s]["date"])}"\n'
        for s in order)
    text = open("index.md", encoding="utf-8").read()
    new = re.sub(r"^recent:\n(?:(?:[ \t].*)?\n)*", lambda m: block, text, count=1, flags=re.M)
    if new != text:
        open("index.md", "w", encoding="utf-8").write(new)
        return True
    return False


def rebuild_all_entries(entries, recent):
    """Write entries.md (the full index linked from the homepage), newest first, nothing excluded."""
    prev = {u: i for i, (u, _) in enumerate(recent)}
    order = sorted(entries, key=lambda s: (-entries[s]["date"].toordinal(), prev.get(s, 10**6), entries[s]["title"]))
    notes = "".join(f'  - title: {q(entries[s]["title"])}\n    url: "/{s}/"\n    date: "{short(entries[s]["date"])}"\n'
                    for s in order)
    text = ('---\nlayout: topic\ntitle: "All entries"\ntopic: "all entries"\npermalink: /entries/\n'
            'description: "Every entry, most recently updated first."\nnotes:\n' + notes + '---\n')
    old = open("entries.md", encoding="utf-8").read() if os.path.exists("entries.md") else None
    if text != old:
        open("entries.md", "w", encoding="utf-8").write(text)
        return True
    return False


def write_topic_page(topic, family, siblings, members, entries):
    path = f"topics/{slugify(topic)}.md"
    notes = "notes:\n" + "".join(
        f'  - title: {q(entries[s]["title"])}\n    url: "/{s}/"\n    date: "{short(entries[s]["date"])}"\n'
        for s in members)
    if os.path.exists(path):
        fm, body = split_fm(open(path, encoding="utf-8").read())
        fm = re.sub(r"^notes:\n(?:[ \t].*\n?)*", "", fm + "\n", flags=re.M).rstrip("\n")
        lines = fm.split("\n")
        at = next((i for i, l in enumerate(lines) if l.startswith("see_also:")), len(lines))
        new = "\n".join(lines[:at]) + "\n" + notes + "\n".join(lines[at:])
        text = "---\n" + new.rstrip("\n") + "\n---\n" + body
    else:
        title = topic[:1].upper() + topic[1:] if family != "Figures" else " ".join(
            w if w == "of" else w[:1].upper() + w[1:] for w in topic.split(" "))
        see = "" if family == "Figures" else "see_also:\n" + "".join(
            f'  - title: {q(x)}\n    url: "/topics/{slugify(x)}/"\n' for x in siblings)
        text = (f"---\nlayout: topic\ntitle: {q(title)}\ntopic: {q(topic)}\n"
                f"permalink: /topics/{slugify(topic)}/\n{notes}{see}---\n")
    old = open(path, encoding="utf-8").read() if os.path.exists(path) else None
    if text != old:
        open(path, "w", encoding="utf-8").write(text)
        return True
    return False


def main():
    check_only = "--check" in sys.argv
    errors, warnings = [], []
    entries = load_entries(errors)
    groups, recent, exclude = load_taxonomy()
    family = {t: g for g, ts in groups for t in ts}

    ids = {}
    for s, e in entries.items():
        if not e["title"]:
            errors.append(f"{s}: missing title")
        if e["date"] is None:
            errors.append(f"{s}: unreadable updated date {e['updated']!r}")
        ids.setdefault(e["id"], []).append(s)
        if not e["topics"]:
            errors.append(f"{s}: no topics")
        for t in e["topics"]:
            if t not in family:
                errors.append(f"{s}: topic '{t}' is not in index.md topic_groups")
        if not e["related"]:
            errors.append(f"{s}: no related entries")
        elif not 3 <= len(e["related"]) <= 5:
            warnings.append(f"{s}: {len(e['related'])} related entries (aim for 3-5)")
        for r in e["related"]:
            if r == s:
                errors.append(f"{s}: related list links to itself")
            elif r not in entries:
                errors.append(f"{s}: related entry '/{r}/' does not exist")
    for i, ss in ids.items():
        if len(ss) > 1:
            errors.append(f"duplicate id {i}: {', '.join(ss)}")

    inbound = {s: 0 for s in entries}
    for e in entries.values():
        for r in e["related"]:
            if r in inbound:
                inbound[r] += 1
    for s, n in inbound.items():
        if n == 0:
            warnings.append(f"{s}: no other entry lists it as related")
    one_way = sum(1 for s, e in entries.items() for r in e["related"]
                  if r in entries and s not in entries[r]["related"])

    counts = {t: 0 for t in family}
    for e in entries.values():
        for t in e["topics"]:
            if t in counts:
                counts[t] += 1
    for t, n in counts.items():
        if n == 0:
            warnings.append(f"topic '{t}' has no entries")
        elif n < 3 and family[t] != "Figures":
            warnings.append(f"topic '{t}' has only {n} entr{'y' if n == 1 else 'ies'}")

    recent_problems = []
    seen = set()
    listed = {u for u, _ in recent}
    for u, d in recent:
        if u in seen:
            recent_problems.append(f"recent list: '/{u}/' appears more than once")
        seen.add(u)
        if u not in entries:
            recent_problems.append(f"recent list: '/{u}/' does not exist")
        elif entries[u]["date"] and d != short(entries[u]["date"]):
            recent_problems.append(f"recent list: '/{u}/' shows {d}, entry updated {short(entries[u]['date'])}")
    for s in entries:
        if s not in listed and s not in exclude:
            recent_problems.append(f"recent list: '/{s}/' is missing")
    for x in exclude:
        if x not in entries:
            warnings.append(f"recent_exclude: '/{x}/' does not exist")
    dated = [entries[u]["date"] for u, _ in recent if u in entries and entries[u]["date"]]
    if any(a < b for a, b in zip(dated, dated[1:])):
        recent_problems.append("recent list is not in newest-first order")
    if check_only:
        errors.extend(recent_problems)

    stale_pages = [p for p in glob.glob("topics/*.md")
                   if os.path.basename(p)[:-3] not in {slugify(t) for t in family}]
    for p in stale_pages:
        warnings.append(f"{p}: topic page for a topic not in the taxonomy")

    changed = []
    if not check_only and not errors:
        for g, ts in groups:
            for t in ts:
                members = sorted((s for s, e in entries.items() if t in e["topics"]),
                                 key=lambda s: (-entries[s]["date"].toordinal(),
                                                -entries[s]["cdate"].toordinal(), entries[s]["title"]))
                if write_topic_page(t, g, [x for x in ts if x != t], members, entries):
                    changed.append(t)
        if rebuild_all_entries(entries, recent):
            changed.append("all-entries index")
        if rebuild_recent(entries, recent, exclude):
            changed.append("recent list (" + "; ".join(recent_problems) + ")" if recent_problems else "recent list")
    elif check_only:
        for g, ts in groups:
            for t in ts:
                path = f"topics/{slugify(t)}.md"
                if not os.path.exists(path):
                    errors.append(f"missing topic page {path}")
        if not os.path.exists("entries.md"):
            errors.append("missing entries.md (all-entries index)")

    print(f"{len(entries)} entries, {len(family)} topics, "
          f"{sum(len(e['related']) for e in entries.values())} related links ({one_way} one-way)")
    if changed:
        print("rebuilt: " + ", ".join(changed))
    for w in warnings:
        print("WARN  " + w)
    for e in errors:
        print("ERROR " + e)
    if errors:
        print("Fix the errors above; topic pages were not rebuilt." if not check_only else "")
        sys.exit(1)
    print("OK")


if __name__ == "__main__":
    main()
