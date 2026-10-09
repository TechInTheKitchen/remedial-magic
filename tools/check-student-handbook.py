from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

path = Path(__file__).resolve().parents[1] / 'Handouts' / 'Prisms Student Handbook 5x7.docx'
doc = Document(path)
s = doc.sections[0]
assert round(s.page_width / 914400, 3) == 5
assert round(s.page_height / 914400, 3) == 7
breaks = [el for el in doc.element.iter(qn('w:br')) if el.get(qn('w:type')) == 'page']
assert len(breaks) == 3
cover = []
for child in doc.element.body:
    if any(el.get(qn('w:type')) == 'page' for el in child.iter(qn('w:br'))):
        break
    cover.extend(el.text or '' for el in child.iter(qn('w:t')))
cover_text = ' '.join(cover)
for field in ['Student name', 'Specialty school', 'Background knack', 'Outstanding business', 'Level', 'Ordinary draw', 'Specialty draw', 'COMPOSURE', 'Current pouch total']:
    assert field in cover_text, field
for table in doc.tables:
    assert round(sum(col.width for col in table.columns) / 914400, 2) <= 4.24
print('Validated DOCX: 5 x 7 inch pages; three explicit page breaks; all core stats on cover; tables within margins.')
print('Actual pagination remains unverified without a document renderer.')
