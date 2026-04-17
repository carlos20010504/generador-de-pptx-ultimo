"""Extract color info from Socya template PPTX"""
import zipfile, re, sys

PPTX = "Plantilla_Presentacion_Socya (1) (1).pptx"
sys.stdout.reconfigure(encoding='utf-8')

colors_found = set()

with zipfile.ZipFile(PPTX) as z:
    names = z.namelist()
    themes = [n for n in names if 'theme' in n and n.endswith('.xml')]
    for t in themes:
        xml = z.read(t).decode('utf-8', errors='replace')
        for m in re.finditer(r'srgbClr val="([0-9A-Fa-f]{6})"', xml):
            colors_found.add(m.group(1).upper())
    
    slides = sorted([n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)])
    slide1_xml = z.read(slides[0]).decode('utf-8', errors='replace') if slides else ''
    slide_last_xml = z.read(slides[-1]).decode('utf-8', errors='replace') if slides else ''
    
    for xml in [slide1_xml, slide_last_xml]:
        for m in re.finditer(r'srgbClr val="([0-9A-Fa-f]{6})"', xml):
            colors_found.add(m.group(1).upper())

    print("=== ALL COLORS FOUND ===")
    for c in sorted(colors_found):
        print(f"  #{c}")

    # Extract theme colors specifically
    if themes:
        xml = z.read(themes[0]).decode('utf-8', errors='replace')
        # dk1, lt1, dk2, lt2, accent1..6
        for tag in ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6']:
            m = re.search(rf'<a:{tag}[^>]*>.*?srgbClr val="([0-9A-Fa-f]{{6}})"', xml, re.DOTALL)
            m2 = re.search(rf'<a:{tag}[^>]*>.*?sysClr[^/]*/>', xml, re.DOTALL)
            if m:
                print(f"  {tag}: #{m.group(1).upper()}")
            elif m2:
                print(f"  {tag}: (system color)")

    # Print slide 1 color rects for header/cover
    print("\n=== SLIDE 1 COLORS (cover) ===")
    for m in re.finditer(r'srgbClr val="([0-9A-Fa-f]{6})"', slide1_xml):
        print(f"  #{m.group(1).upper()}")

    print(f"\nTotal slides: {len(slides)}")
