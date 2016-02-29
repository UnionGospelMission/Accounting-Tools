from distutils.core import setup
import py2exe,os,shutil,sys,zipfile,json
windowed=True
scriptname=False
print 'usage: compiler.py py2exe scriptname=[yourscript.py] [--consoled] [--includes="include1,include2] [--singlefile]"'
includes = []
kwargs = {'options':{'py2exe':{}}}

windowed = 'windows'
for i in sys.argv[:]:
	if "scriptname=" in i:
		scriptname=sys.argv.pop(sys.argv.index(i)).split('=')[1]
	if '--consoled' in i:
		windowed='console'
		a = sys.argv.pop(sys.argv.index(i))
	if '--includes=' in i:
		includes = sys.argv.pop(sys.argv.index(i)).split('=')[1].replace('"','').split(',')
	if '--singlefile' in i:
		kwargs['options']['py2exe'].update({'bundle_files': 1, 'compressed': True})
		sys.argv.pop(sys.argv.index(i))

kwargs['options']['py2exe'].update({'includes': includes, 'dll_excludes': ['w9xpopen.exe','MSVCR71.dll']})
kwargs[windowed] = [{'script': scriptname}]
kwargs['zipfile'] = None

if scriptname:
	if os.path.isdir('build'):
		os.system('rmdir build /s /q')
	if os.path.isdir('dist'):
		os.system('rmdir dist /s /q')
	if os.path.isdir(scriptname.split('.')[0]):
		os.system('rmdir %s /s /q'%scriptname.split('.')[0])
	mylist = [{"script":scriptname,"icon_resources": [(1, "icon.ico")], "includes":includes}]
	setup(**kwargs)
	
	shutil.move("dist",scriptname.split('.')[0])

'''
setup(
    options = {'py2exe': {'bundle_files': 1, 'compressed': True, 'includes': includes,
                'dll_excludes': ['w9xpopen.exe','MSVCR71.dll']}},
    console = [{'script': scriptname}],
    zipfile = None,
)
'''
