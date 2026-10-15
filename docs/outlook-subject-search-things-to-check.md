# Outlook subject search: things to check

## What is documented

The equality-based exact search path removes leading `RE:` and `FW:` from the
search text, asks Outlook for `[Subject] = "base subject"`, then compares the
returned subjects in Python after removing those prefixes. Python cannot
recover a message that Outlook excluded from the restriction.

[Microsoft's subject filtering example](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/filtering-items-using-a-string-comparison)
shows that `[Subject] = 'cat'` returns both `cat` and `RE: cat`. It also says
standard `FW: ` prefixes are ignored. This supports the fast equality lookup
for ordinary replies and forwards.

## What still needs a live Outlook check

1. **No space after the prefix.** The [Exchange subject specification](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmail/285b2511-9f7b-45d9-994e-98b097ab0ae8)
   and [PST specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/ba6398a3-82ba-46aa-bbe9-3eb28917c271)
   allow zero spaces after the colon. The [Outlook MAPI property documentation](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagnormalizedsubject-canonical-property)
   describes a colon followed by a space when calculating a missing prefix.
   We have not confirmed whether `Items.Restrict` returns `RE:cat` for a lookup
   of `cat` in each store type.
2. **Stacked prefixes.** The [normalized subject property](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagnormalizedsubject-canonical-property)
   describes removing a recorded prefix, but does not promise repeated removal.
   We have not confirmed whether a lookup of `cat` returns `FW: RE: cat` or
   `FW:RE: cat`.
3. **Actual lookup time.** Unit tests use fake Outlook items, so they cannot
   measure the time spent sorting, applying date and subject restrictions,
   enumerating returned items, or reading message bodies. We have not timed
   the restored equality lookup in Windows Outlook.

The subject property specifications above describe how subjects are stored;
they do not directly establish what `Items.Restrict` returns for the unusual
prefix forms.

## How to check later

In a representative Outlook folder, place messages with subjects `cat`,
`RE: cat`, `FW: cat`, `RE:cat`, `FW: RE: cat`, and `FW:RE: cat`. Run an exact
search for both `cat` and `RE: cat`; record the returned subjects. Compare them
with the items from `Items.Restrict('[Subject] = "cat"')`. Record the Outlook
version and whether the folder is backed by Exchange or a PST file.

For performance, time sorting, each `Restrict` call, and enumeration separately
on a folder similar in size to the one that felt slow. Count items returned by
Outlook as well as final Python matches; a small final result does not show how
many items were examined before filtering.
