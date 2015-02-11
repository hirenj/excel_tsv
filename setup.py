from distutils.core import setup
import py2exe

setup(
	console=['extract_excel_sheet','write_excel'],
	options={
		'py2exe' : {
			'packages' : [ 'xlrd','xlwt' ],
		}
	},
	install_requires=[
		'xlrd',
		'xlwt'
	],
)