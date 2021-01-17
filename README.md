# The So-What
*consultify* allows users to turn analyses on python into PowerPoint slides, following the conventions of standard management consulting slides.

# Installation
```
pip install consultify
```

# Usage
```
from consultify import consultify

prs = consultify.make_deck()

consultify.add_slide(prs, slide_title='Title', image_filepath='sample.jpg', slide_text=
"""Bullet 1
Bullet 2
Bullet 3""",)

consultify.add_marvin_table_slide(prs, df, slide_title = 'Title')

consultify.save_deck(prs, filepath='./210117 SteerCo Deck.pptx')
```
See attached Jupyter Notebook for additional examples.

# #plsfix
Email ryu@mba2021.hbs.edu for feedback.