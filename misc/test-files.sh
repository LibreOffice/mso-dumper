#! /bin/sh

test_dir=$1

if [ ! -d $test_dir ]
then
    echo "Usage: test-files.sh TEST_DIR"
    exit 1
fi

for x in `find $test_dir -name \*.xls`; do
    (python xls-dump.py $x | grep -v "rror inter" > /dev/null) || echo "Flat dump failed for" $x
    python xls-dump.py --dump-mode=xml $x > /dev/null || echo "Xml dump failed for" $x
    python xls-dump.py --dump-mode=cxml $x > /dev/null || echo "CXml dump failed for" $x
done
