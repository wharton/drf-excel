from setuptools import setup, find_packages

with open("README.md") as f:
    long_description = f.read()

setup(
    name="drf-renderer-xlsx",
    description="Django REST Framework renderer for spreadsheet (xlsx) files.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Timothy Allen",
    author_email="tallen@wharton.upenn.edu",
    url="https://github.com/wharton/drf-renderer-xlsx",
    include_package_data=True,
    packages=find_packages(),
    zip_safe=False,
    install_requires=["djangorestframework>=3.6", "openpyxl>=2.4"],
    setup_requires=["setuptools_scm"],
    use_scm_version=True,
    classifiers=[
        "Development Status :: 4 - Beta",
        "Environment :: Web Environment",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: BSD License",
        "Operating System :: OS Independent",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3 :: Only",
        "Framework :: Django",
        "Framework :: Django :: 2.2",
        "Framework :: Django :: 3.0",
        "Framework :: Django :: 3.1",
        "Framework :: Django :: 3.2",
        "Topic :: Internet :: WWW/HTTP",
        "Topic :: Internet :: WWW/HTTP :: Dynamic Content",
    ],
)
