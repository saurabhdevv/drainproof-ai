# DrainProof AI - prototype (synthetic demo data)

- dashboard.html : open in a browser (needs the img/ folder next to it). Views: #map, #tasks, #verify, #log
- backend.py     : trains the risk model, simulates drain photos, runs the agent (prioritise, assign, verify). Writes data.json
- build_and_shoot.py : rebuilds dashboard.html and the screenshots in shots/
- dashboard.template.html : dashboard source
- build_deck.js  : generates DrainProof_AI_Pitch.pptx (pptxgenjs)
- shots/         : the 4 screenshots used in the deck
- img/           : synthetic before/after drain photos

Run: pip install scikit-learn pandas pillow playwright ; python3 backend.py ; python3 build_and_shoot.py
