""" The IniParams dictionary is simply a regular dictionary.
We just have it here because this is a good place to put
advanced options for the Heat Source module.

The idea here is that there are certain options
which we should include in the interface for daily
use, and certain options that we should hide in a place
where, to put it bluntly, people won't screw with them
unless they're smart enough to know what they're doing.

Of course, you can't program against stupidity, but there you are.

Anyway, this is just a convenient place to hold them, where
they can be found at a later time. Of course, one important thing
to remember is that since many classes are built with the psyco
optimize switch, this module should be imported before any other
modules. Including the import in the Dieties.__init__.py module.

Good thing that very important caveat is buried deeply in these
notes where no-one will ever read it.

To turn of psyco alltogether, edit heatsource.__init__.py
"""


IniParams = {"psyco": ('Dictionaries',
                        'BigRedButton',
                        'StreamNode',
                        'PyHeatsource',
                        'ExcelDocument',
                        'ExcelInterface',
                        'IniParamsDiety',
                        'ChronosDiety',
                        'Output'),
             # Run the routines in PyHeatsource.py instead of
             # the C module.
             "run_in_python": False,
             }
