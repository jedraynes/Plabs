# Plabs

A part of the financial diligence process is conveying financial results. In my company, this is done through a final output summarizing findings as defined by our engagement scope.

I noticed a pain point in the process that occurred on every deal: when presenting annual income statement and balance sheet trends, I found myself repeating basic words in the write-ups for each financial statement line item (often truncated as FSLI).

To solve this, I created a Python script dubbed Plabs, an acronym for "profit/loss and balance sheet", the two statements trended in the diligence process. The script takes an Excel input, one already created as part of the data modeling part of processing trial balances, and drafts the write-ups based on the accounts in the flat file. This output can then be pasted, as unformatted text, into the working copy of our final report. Then, the user can make final edits to customize the write-ups as needed for each deal. At a minimum, the script saves user time and fixes a small, but repetitive task I've noticed on every engagement.

**Libraries:** Pandas, Sys

---
Feel free to check out my personal blog at [jedraynes.com](https://www.jedraynes.com). There, you can find ways to contact me via a contact form or over [LinkedIn](https://www.linkedin.com/in/jedraynes/).

jedraynes
