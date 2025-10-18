# Make local tests a real package to avoid shadowing by site-packages 'tests'
# This ensures imports like `from tests.conftest import stub_xlwt` resolve here.