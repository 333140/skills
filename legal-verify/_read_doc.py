import sys
sys.path.insert(0, r'g:\TRAE\04-skill\.trae\skills\legal-verify\scripts')
from parse_citations import read_file

text = read_file(r'g:\TRAE\04-skill\.trae\skills\一.docx')
print(text[:3000])
print("...")
print(f"Total length: {len(text)}")
