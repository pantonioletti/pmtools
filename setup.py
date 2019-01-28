from setuptools import setup

setup(name="JDAPMTools",
      version="1.0",
      scripts=['UsingPandas.py'],
      install_requires=['pytz', 'numpy', 'six', 'python-dateutil', 'pandas', 'jdcal', 'et-xmlfile', 'openpyxl', 'xlrd'],
      author='Pablo Antonioletti',
      author_email='pablo.antonioletti@jda.com'

)
