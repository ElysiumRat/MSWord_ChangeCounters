# MSWord_ChangeCounters
A collection of attempts to write VBA macros for Microsoft Word to count the numbers of changes by specific authors. While this information can be gained easily from Word, the need for macros to count these were for work-related calculations (not included in these macros) about who exactly did how much, so as to save time in the workflow by not having to calculate these manually (again, it's perfectly possible to do so, but having Word do it automatically would speed up certain processes).

All macros work, but as soon as a document becomes large, they become too slow to be of any use. The main reason I am sharing them is because when I was looking for a solution to this exact problem, none were immediately available, so perhaps in sharing them, people will be able to use them (or see that they're not useable for certain projects). Brief descriptions of each attempt follow.

UPDATE: I have created [a working version of the idea](https://github.com/ElysiumRat/MSWord_AuthorCheck) that I decided to upload separately for the sake of clarity. These failed attempts will remain accessible for posterity, as perhaps others may be able to learn from these failed attempts.

## AuthorArrayCheck
My first attempt. This creates arrays of authors and the number of changes and comments each author has contributed to the document, before giving a complete list of who did how much of each. It works, but it's too slow to be useable in larger documents (tests on a document with 1600+ changes were abandoned after 15 minutes of Word being unresponsive).

## ChangeCheck
An attempt to speed up the AuthorArrayCheck macro, this one asks you to input an author name, then checks all changes and comments in document for who created them, tallying if the author name matches the inputted name, before giving a list of how much of each were done by that specific person. Again, it works, but it doesn't seem to speed up the AuthorArrayCheck process much (tests on the same large document mentioned above produced much the same result).

## SelectionCheck
A version of ChangeCheck, this asks you to input an author name, and checks all changes and comments, but only for a selection of text, rather than throughout the entire document. This is probably the one least likely to hang, but the usefulness of this is extremely limited, ultimately.
