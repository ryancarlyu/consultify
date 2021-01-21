import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()
with open('requirements.txt') as f:
    required = f.read().splitlines()

setuptools.setup(
    name="consultify", # Replace with your own username
    version="0.0.2.2",
    author="Ryan Yu",
    author_email="ryu@mba2021.hbs.edu",
    description="Turns analyses on Python into management consulting-style PowerPoint slides.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ryancarlyu/consultify",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    install_requires=required
)