from setuptools import setup, find_packages

setup(
    name="pptx_to_pdf",
    version="0.1.0",
    author="Your Name",
    description="A Python package to convert PPTX files to PDF using PowerPoint",
    packages=find_packages(),
    install_requires=["pillow", "pywin32"],
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: Microsoft :: Windows",
    ],
)
