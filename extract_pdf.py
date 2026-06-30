import fitz

doc = fitz.open(r"c:\Users\shiba\OneDrive\Desktop\Cognexa AI-Autonomous-Multi-Agent-Research-System\ACE_Presentation_Shiba.pdf")
with open(r"c:\Users\shiba\OneDrive\Desktop\Cognexa AI-Autonomous-Multi-Agent-Research-System\pdf_output.txt", "w", encoding="utf-8") as f:
    for i, page in enumerate(doc):
        f.write(f"\n=== SLIDE {i+1} ===\n")
        f.write(page.get_text())
doc.close()
print("Done! Saved to pdf_output.txt")
