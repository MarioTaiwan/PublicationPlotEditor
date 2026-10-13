# Static checks (no MATLAB runtime)
from pathlib import Path
import json
from xml.etree import ElementTree
p=Path(__file__).parent
src=(p/'PublicationPlotEditor.m').read_text()
master=json.loads((p/'StandardScientific.json').read_text())
ElementTree.parse(p/'Example_Corel_Palette.xml')
assert master['version']==7
assert set(master['export']['FontFactors'])=={'XTick','YTick','XLabel','YLabel'}
assert 'function openEMFFontCalibration(~,~)' in src
assert src.count('try captureVisibleFontSizes();catch err') == 3
print('PASS static source/JSON/XML checks')
