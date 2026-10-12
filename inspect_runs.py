import docx

doc_path = r"C:\Users\harle\Downloads\Offer Letter Harley.docx"
doc = docx.Document(doc_path)

for i in [0, 2, 3, 4, 5, 9, 11, 12, 13, 24]:
    p = doc.paragraphs[i]
    print(f"=== P{i} ===")
    print("Full text:", repr(p.text))
    for j, r in enumerate(p.runs):
        print(f"  r{j}: {repr(r.text)} | bold={r.bold} | italic={r.italic} | font={r.font.name} | size={r.font.size.pt if r.font.size else None} | color={r.font.color.rgb if r.font.color else None}")

