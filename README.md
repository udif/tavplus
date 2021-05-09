This script is used to manage your shopping database in https://tavplus.mltp.co.il/ gift cards.
Yopu need to feed it with an XLSX file containing a list of card IDs on column A.
Each card is queried for its balance and a transaction list. Results are kept in a pickle file,
so that subsequent queries will only query cards with positive balance. Cards with zero balance (empty)
are kept in the pickle file but not queried again.

The output of this iscript is a different XLSX file with a detailed transaction list
that can be used to analyze expenses over time.

