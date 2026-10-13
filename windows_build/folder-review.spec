from pathlib import Path
from zipfile import ZipFile
import os

root = Path(SPECPATH).parent
sevenzip = Path(os.environ['FOLDER_REVIEW_7ZIP_DIR'])
testing = os.environ.get('FOLDER_REVIEW_BUILD_TESTS') == '1'
name = 'folder-review-tests' if testing else 'folder-review'
entry = root / 'windows_build' / ('test_entry.py' if testing else 'entry.py')
data = [(str(sevenzip / 'License.txt'), 'licenses/7zip')]
if testing:
    data.append((str(root / 'tests' / 'fixtures'), 'fixtures'))

a = Analysis(
    [str(entry)],
    pathex=[str(root), str(root / 'tests')],
    binaries=[(str(sevenzip / '7z.exe'), '.'), (str(sevenzip / '7z.dll'), '.')],
    datas=data,
    hiddenimports=[],
    # These are builder-only packages, absent from a clean uv.lock runtime.
    # Keep all runtime codecs, native libraries, assertions and docstrings.
    excludes=['setuptools', 'pkg_resources', 'packaging'],
    optimize=0,
)
# PyInstaller's module graph can produce a different standard-library ZIP
# entry order between processes. Keep every entry and its metadata unchanged;
# sort only the container order before the standard PyInstaller packaging step.
for dest_name, source_name, _ in a.datas:
    if dest_name == 'base_library.zip':
        library = Path(source_name)
        ordered = library.with_suffix('.ordered.zip')
        with ZipFile(library) as source, ZipFile(ordered, 'w') as target:
            target.comment = source.comment
            for info in sorted(source.infolist(), key=lambda item: item.filename):
                target.writestr(info, source.read(info.filename))
        ordered.replace(library)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name=name,
    debug=False,
    strip=False,
    upx=False,
    console=True,
    version=str(root / 'windows_build' / 'version.txt'),
)
