check:
	cd test/doc && ./test.py
	cd test/emf && ./test.py
	pep8 --ignore=E501 msodumper/{binarystream,msometa}.py
	pep8 --ignore=E501 doc-dump.py msodumper/doc{record,sprm,stream}.py test/doc/test.py
	pep8 --ignore=E501 emf-dump.py msodumper/{emf,wmf}record.py
	pep8 --ignore=E501 vsd-dump.py msodumper/vsdstream.py test/vsd-test.py
	pycodestyle --ignore=E501 swlaycache-dump.py msodumper/swlaycacherecord.py
