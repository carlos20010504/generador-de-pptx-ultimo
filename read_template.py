import markitdown
md = markitdown.MarkItDown()
r = md.convert("Plantilla_Presentacion_Socya (1) (1).pptx")
# Print first slides (slide 1 and 2 usually have the design info)
text = r.text_content
# Print first 3000 chars to see cover design
print("=== FIRST 3000 CHARS ===")
print(text[:3000])
