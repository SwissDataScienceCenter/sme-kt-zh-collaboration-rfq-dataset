---
status: accepted
---

# Write dataset documents in Swiss German, framework documents in English

All Case and Company world documents are written directly in Swiss German (no ß), in one language only, with no translation step. Swiss trade terms such as Regiearbeiten, bauseits, Eventualposition, Skonto and Ausmass have no stable English equivalent, and translating at the end would let line item names drift apart across documents that reference each other. Framework documents (glossary, ADRs, README, generation specs) stay in English, and the glossary carries the German term for each concept.
