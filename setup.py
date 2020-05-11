from setuptools import setup

setup(
    name="compare-referrer-rcti",
    version="1.0.0",
    description="Project to compare commission files of Infynity and Loankit",
    long_description="README",
    long_description_content_type="text/markdown",
    url="https://github.com/petrosschilling/commission-comparer-infynity",
    author="Petros Schutz Schilling",
    author_email="petros.schilling@loankit.com.au",
    license="MIT",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python",
        "Programming Language :: Python :: 2",
        "Programming Language :: Python :: 3",
    ],
    packages=["src"],
    include_package_data=True,
    install_requires=[
        "beautifulsoup4",
        "click",
        "flake8",
        "xlsxwriter",
        "pandas",
        "xlrd"
    ],
    entry_points={"console_scripts": ["infynity=cli:main"]},
)
