from setuptools import setup, find_packages

setup(
    name='drf-renderer-xlsx',
    version="0.2.2",
    description='Django REST Framework renderer for spreadsheet (xlsx) files.',
    long_description='',
    author='Timothy Allen',
    author_email='tallen@wharton.upenn.edu',
    url='https://github.com/wharton/drf-renderer-xlsx',
    include_package_data=True,
    packages=find_packages(),
    zip_safe=False,
    install_requires=[
        'djangorestframework>=3.6',
        'openpyxl>=2.4',
    ],
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Environment :: Web Environment',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Framework :: Django',
        'Topic :: Internet :: WWW/HTTP',
        'Topic :: Internet :: WWW/HTTP :: Dynamic Content',
    ],
)
