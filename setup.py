from distutils.core import setup
import py2exe

setup(
	console=['extract_excel_sheet','write_excel'],
	zipfile=None,
	options={
		'py2exe' : {
			'packages' : [ 'xlrd','xlwt' ],
			'dll_excludes':['w9xpopen.exe'],
			"compressed" : 1,
			"optimize" : 2,
			"bundle_files" : 1
		}
	},
	install_requires=[
		'xlrd',
		'xlwt'
	],
)