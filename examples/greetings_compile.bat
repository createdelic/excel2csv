.\..\dist\excel2csv.exe --trim --unify-quotemarks -i greetings.xlsx -o out\Greetings_EN -p paths_greetings.json -f 1 Yes No -c 2 3

.\..\dist\excel2csv.exe --trim --unify-quotemarks -i greetings.xlsx -o out\Greetings_JP -p paths_greetings.json -f 1 Yes No -c 2 4

.\..\dist\excel2singlecsv.exe --quotechar ' --delimiter , --trim --unify-quotemarks -i greetings.xlsx -o out\GreetingsAll.csv -f 1 Yes No -c 2 3 4