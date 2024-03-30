# Annual overview


## Missing Photos 

```dataview
TABLE WITHOUT ID
  link(RegistrationNumber) AS RegistrationNumber,
  DescriptionOfGift AS Gift,
  ("![img|100](" + LinkToPhoto +")") as Photo
FROM "gifts" 
WHERE LinkToPhoto
``` 

## Missing Photos 

```dataview
TABLE WITHOUT ID
  link(RegistrationNumber) AS RegistrationNumber,
  NameOfMEP AS Name,
  link(LinkToPhoto) AS Photo
FROM "gifts" 
WHERE !LinkToPhoto
``` 

## ByGifts 

```dataview
TABLE WITHOUT ID
  NameOfMEP AS Name,
  giftCount as Gifts
FROM "gifts" 
GROUP BY NameOfMEP
FLATTEN length(rows) as giftCount
SORT giftCount DESC
``` 

# Background

## MEPS Gifts Registry
The European Parliament published the gifts register as PDF online.
PDF does not provide very good options for data analytics.

- [VIII term - Gifts](https://www.europarl.europa.eu/pdf/meps/gifts_register_8.pdf)
- [IX term - Gifts](https://www.europarl.europa.eu/pdf/meps/gifts_register_9.pdf)

## PDF to XLS conversion
Conveniently, Adobe provides a [Convert PDF to Excel](https://www.adobe.com/acrobat/online/pdf-to-excel.html).
However, the conversion does not take into account the link in the document.
There are many other options[^others], none of which seem to support extracting links.



[^others]: [How to copy a table from PDF to Excel: 8 quick methods](https://nanonets.com/blog/copy-tables-from-pdfs-excel/)