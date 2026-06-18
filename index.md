---
layout: default
title: Home
---

# Moore's Treasury

A commonplace book, in the old sense: a place where things worth remembering are kept. Topics I'm reading, thinking about, or circling. Notes toward works in progress and works in glacial progress. Quotations, references, half-formed arguments.

Organized by subject, not by date. Updated when something new belongs.

---

## Notebooks

{% assign sorted = site.notebooks | sort: "title" %}
{% for nb in sorted %}
- [{{ nb.title }}]({{ nb.url }})
{% endfor %}
