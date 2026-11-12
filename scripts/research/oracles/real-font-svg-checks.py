"""Optional actual PNG corruption controls for the named SVG raster oracle."""
import sys, os, shutil, tempfile, importlib.util, json
from PIL import Image

directory, count = sys.argv[1:]
spec = importlib.util.spec_from_file_location('real_font_svg', os.path.join(os.path.dirname(__file__), 'real-font-svg.py'))
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)
positive = oracle.inspect(directory, count)
rejected = []
with tempfile.TemporaryDirectory(prefix='svg-font-negative-', dir='.tmp') as temporary:
    for index in range(int(count)):
        for kind in ('inline', 'image'):
            file = kind + '-' + str(index) + '.png'
            shutil.copyfile(os.path.join(directory, file), os.path.join(temporary, file))
    for kind in ('invisible', 'changed-pixel', 'wrong-size'):
        for name in ('inline-0.png', 'image-0.png'):
            shutil.copyfile(os.path.join(directory, name), os.path.join(temporary, name))
        file = os.path.join(temporary, 'image-0.png')
        if kind == 'invisible':
            Image.new('RGB', (550, 120), 'white').save(file)
        elif kind == 'changed-pixel':
            image = Image.open(file).convert('RGB')
            image.putpixel((0, 0), (0, 0, 0))
            image.save(file)
        else:
            Image.open(file).crop((0, 0, 549, 120)).save(file)
        try:
            oracle.inspect(temporary, count)
        except AssertionError:
            rejected.append(kind)
        else:
            raise AssertionError('Oracle accepted actual PNG corruption: ' + kind)
print(json.dumps({'positivePairs': len(positive['pairs']), 'negativeControlsRejected': rejected}))
