# Excel 13Fs

## About
This exports 13F data from [aggregate13Fs](https://github.com/cpitzak/aggregate13Fs) to excel for easy filtering and pivoting.

**Disclaimer: I take no responsibility for the correctness or use of this tool. Use it at your own risk.**

## Prerequisites
You need to have the following installed:

- [MongoDB](http://www.mongodb.org)
- [Aggregate13Fs](https://github.com/cpitzak/aggregate13Fs)
- [PyMongo](https://pypi.python.org/pypi/pymongo/2.6.3)
- [ystockquote](https://pypi.python.org/pypi/ystockquote/0.2.4)
- [OpenPyXL](https://pypi.python.org/pypi/openpyxl/1.6.2)

You must have the following running:

- [MongoDB](http://www.mongodb.org)

## Run
If you haven't run [Aggregate 13Fs](https://github.com/cpitzak/aggregate13Fs) in order to populate your mongodb with aggregated 13F form data then first run [you must be online]:

   	$ java -jar build/13FAggregator.jar -update

To export results to excel run [you can be offline]:

	$ python export13Fs gurus13Fdata.xlsx
	
To export results to excel with current share price column run [you must be online]:

	$ python export13Fs -s gurus13Fdata.xlsx
