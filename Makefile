check:
	cd test/doc && ./test.py
	cd test/emf && ./test.py
	pycodestyle --ignore=E501 msodumper/binarystream.py
	pycodestyle --ignore=E501 msodumper/msometa.py
	pycodestyle --ignore=E501 doc-dump.py msodumper/doc*.py test/doc/test.py
	pycodestyle --ignore=E501 emf-dump.py msodumper/*mfrecord.py
	pycodestyle --ignore=E501 vsd-dump.py msodumper/vsdstream.py test/vsd-test.py
	pycodestyle --ignore=E501 swlaycache-dump.py msodumper/swlaycacherecord.py
	pycodestyle --ignore=E501 ole1-dump.py msodumper/ole1record.py
	pycodestyle --ignore=E501 ole2preview-dump.py msodumper/ole2previewrecord.py
	pycodestyle --ignore=E501 convert-enum.py
