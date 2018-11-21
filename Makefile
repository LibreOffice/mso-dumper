check:
	cd test/doc && ./test.py
	cd test/emf && ./test.py
	pycodestyle --ignore=E501 msodumper/{binarystream,msometa}.py
	pycodestyle --ignore=E501 doc-dump.py msodumper/doc{record,sprm,stream}.py test/doc/test.py
	pycodestyle --ignore=E501 emf-dump.py msodumper/{emf,wmf}record.py
	pycodestyle --ignore=E501 vsd-dump.py msodumper/vsdstream.py test/vsd-test.py
	pycodestyle --ignore=E501 swlaycache-dump.py msodumper/swlaycacherecord.py
	pycodestyle --ignore=E501 ole1preview-dump.py msodumper/ole1previewrecord.py
	pycodestyle --ignore=E501 ole2preview-dump.py msodumper/ole2previewrecord.py
