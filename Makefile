check:
	cd test/doc && ./test.py
	pep8 --ignore=E501 doc-dump.py msodumper/doc{dirstream,record}.py
