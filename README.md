# The So-What
*consultify* allows users to turn analyses on Python into PowerPoint slides, following the conventions of standard management consulting slides.

# Installation
```
pip install consultify
```

# Sample Usage
```
# Get sample image from Matplotlib

import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import urllib.request 

urllib.request.urlretrieve('https://matplotlib.org/_images/sphx_glr_bar_stacked_001.png', 'sample.png')
img = mpimg.imread('sample.png')
plt.axis('off')
plt.imshow(img)
```
![Matplotlib sample](https://raw.githubusercontent.com/ryancarlyu/consultify/main/screenshots/sample.png)


```
# Create sample pandas DataFrame

import pandas as pd
from sklearn.datasets import load_iris
data = load_iris()
df = pd.DataFrame(data.data, columns=data.feature_names)
df = df.head(5)
df
```
![Iris DataFrame](https://raw.githubusercontent.com/ryancarlyu/consultify/main/screenshots/iris_dataframe.png)

```
from consultify import consultify

prs = consultify.make_deck()

consultify.add_slide(prs, slide_title='The highest scores were achieved in Game 2', image_filepath='sample.png', textbox_filled_space = 0.25, textbox_font_size = 20, title_font_size = 32, slide_text=
"""Insight # 1

Insight # 2

Insight # 3""")

consultify.add_marvin_table_slide(prs, df, title_font_size = 28, slide_title = 'These are the first five rows of the classic Iris Dataset')

consultify.save_deck(prs, filepath='./210117 SteerCo Deck.pptx')
```
![Output Slide 1](https://raw.githubusercontent.com/ryancarlyu/consultify/main/screenshots/Slide1.png)
![Output Slide 2](https://raw.githubusercontent.com/ryancarlyu/consultify/main/screenshots/Slide2.png)

# #plsfix
Email ryu@mba2021.hbs.edu for feedback.