# exca-history

In this repo we track the historical changes of the Catalhoyuk Excavation database front end.

`.vba` files were extracted as text from the MS Acess Database `.mdb`

Lines that only contain comments and blank lines were removed with:

    for file in *.vba;
    do
      LC_ALL=C sed '/^[[:space:]]*$/d' $file > tmpfile1
      LC_ALL=C sed "/^[[:space:]]\{0,\}\'/ d" tmpfile1 > tmpfile2
      cp tmpfile2 $file
      rm tmpfile*
    done

Files that have 2 lines or less were then removed with:

    wc -l *.vba | awk '{if ($1 <= 2) print}' | awk '{for (i = 2; i < NF; i++) printf $i "\\ "; print $NF}' | xargs rm -f

http://catalhoyuk.com
