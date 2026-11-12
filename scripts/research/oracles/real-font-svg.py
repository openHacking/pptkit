"""Independent PNG visibility and inline/image raster comparison for named SVGs."""
import sys, os, json, hashlib
from PIL import Image, ImageChops

def inspect(directory, count, unembedded_case=None):
    results = []
    for index in range(int(count)):
        inline_file = os.path.join(directory, 'inline-' + str(index) + '.png')
        image_file = os.path.join(directory, 'image-' + str(index) + '.png')
        inline = Image.open(inline_file).convert('RGB')
        image = Image.open(image_file).convert('RGB')
        assert inline.size == image.size == (550, 120)
        def ink(frame):
            pixels = frame.load()
            points = [(x, y) for y in range(frame.height) for x in range(frame.width) if min(pixels[x, y]) < 200]
            return {'pixels': len(points), 'bounds': [min(x for x, y in points), min(y for x, y in points), max(x for x, y in points) + 1, max(y for x, y in points) + 1]} if points else {'pixels': 0, 'bounds': None}
        before, after = ink(inline), ink(image)
        assert before['pixels'] > 100 and after['pixels'] > 100, 'text is not visibly painted'
        difference = ImageChops.difference(inline, image)
        assert difference.getbbox() is None, 'inline/image raster differs for named exact-face control'
        results.append({'case': index, 'inline': before, 'image': after, 'exactRasterMatch': True, 'differentPixels': 0, 'inlinePngSha256': hashlib.sha256(open(inline_file, 'rb').read()).hexdigest(), 'imagePngSha256': hashlib.sha256(open(image_file, 'rb').read()).hexdigest()})
    negative = None
    if unembedded_case is not None:
        source = Image.open(os.path.join(directory, 'inline-' + str(unembedded_case) + '.png')).convert('RGB')
        image = Image.open(os.path.join(directory, 'unembedded.png')).convert('RGB')
        assert source.size == image.size == (550, 120)
        assert ImageChops.difference(source, image).getbbox() is not None, 'unembedded image unexpectedly inherited the document alias face'
        negative = {'case': int(unembedded_case), 'unembeddedImageDiffers': True}
    return {'pairs': results, 'isolatedFontControl': negative}

if __name__ == '__main__':
    print(json.dumps(inspect(*sys.argv[1:])))
