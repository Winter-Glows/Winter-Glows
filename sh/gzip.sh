ls | while read id
do

gzip -c $id > $id.gz

done

