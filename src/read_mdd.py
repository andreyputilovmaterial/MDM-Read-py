"""Read MDD files (IBM SPSS Data Collection / Unicom Intelligence) and dump them to JSON.

The proprietary MDD format is only readable through the COM API exposed by
IBM/Unicom Professional, so opening a document requires Windows, IBM/Unicom
Professional, and ``pywin32`` (``win32com.client``).

The CLI entry point is :func:`entry_point`; the work itself lives in
:class:`MDMDocument`, whose ``read`` method returns a dict suitable to be saved or
sent as JSON.
"""


# import os, time, re, sys
from datetime import datetime, timezone
# from dateutil import tz
import argparse
from pathlib import Path
import re
import json
import traceback, sys




# # import pythoncom
# import win32com.client
# from win32com_adapter import Win32Client
if __name__ == '__main__':
    # run as a program
    from GENERATED._VERSION import _VERSION as script_version
    from win32com_adapter import Win32Client
elif '.' in __name__:
    # package
    from .GENERATED._VERSION import _VERSION as script_version
    from .win32com_adapter import Win32Client
else:
    # included with no parent package
    from GENERATED._VERSION import _VERSION as script_version
    from win32com_adapter import Win32Client




# STDOUT_COLOR_RED = "\033[91m"
STDOUT_COLOR_RED = "\033[31m"
STDOUT_COLOR_RESET = "\033[0m"
STDOUT_COLOR_GREEN = "\033[32m"






CONFIG_LISTCAT_STEPINTO_SL = True



std_col_labels = {
    'name': 'Item name',
    'label': 'Label',
    'attributes': 'Item Attributes',
    'properties': 'Custom properties',
    'translations': 'Translation overlay ',
    'scripting': 'Scripts'
}



def normaize_linebreaks(s):
    """Collapse any sequence of CR/LF characters into a single ``\\n``."""
    return re.sub(r'(?:(?:\r)|(?:\n))+',"\n",s)




# how to use this class
# 1. create an object of type MDMDocument (and pass the path to MDD as a param)
# the MDD is opened as a file system object, so the connection is made to the file (descriptor is acquired)
# also, the third param is to pass config options - which properties to read (for example, just labes, or translations too), which sections, etc...
# I don't have cofig options documented, please read through source code to understand config power
# 2. call "read"
# so the program iterates over all fields
# and returns result as dict that can be represented as json
# that's it, enjoy

# when done, file descriptor is released (this is also happening on exceptions because it is specced in __del__ method)

class MDMDocument:
    """Open and read an MDD document, returning its contents as a JSON-ready dict.

    Usage:
        1. ``doc = MDMDocument(mdd_path)`` to open the document.
        2. ``result = doc.read(config)`` to walk it and produce the report dict.
        3. Use it as a context manager (``with MDMDocument(...) as doc:``); the
           underlying COM document is also released in ``__del__``.
    """

    # a helper error class
    class ErrorItemNotIterable(Exception):
        """Raised from list_categories()"""
        pass

    # constructor: load MDD document here
    # three params:
    # 1. the path to MDD,
    # 2. an optional open method: "open" or "join"
    #    - "open" is simpler/straightforward and usually keeps the normal label
    #      representation in the "Scripts" view, but it opens the file exclusively
    #      (blocking), so it may fail if the file is already in use;
    #    - "join" creates a new empty document and merges the opened document into it;
    #      for a deep, recursive read with explicit contexts it makes little difference,
    #      and it does not need exclusive access, so it is a more stable fallback.
    #    Default is "open"; "join" is equally legitimate.
    # 3. config options (see the entry_point CLI for the documented set)
    def __init__(self,mdd_path,method='open',config=None):
        """Open the MDD document via the MDM COM API.

        Args:
            mdd_path: Path to the MDD file to open.
            method: ``"open"`` or ``"join"``. ``"open"`` is the default and the more
                straightforward mode; ``"join"`` merges the document into a new empty
                one and does not require exclusive access to the file (see the comments
                in the method body for the trade-offs).
            config: Optional read options; defaults are merged for the keys
                ``features``, ``sections`` and ``contexts``.

        Raises:
            FileNotFoundError: If ``mdd_path`` does not exist.
            ValueError: If ``method`` is neither ``"open"`` nor ``"join"``.
        """

        if config is None:
            config = {}

        self.__document = None

        if method=='open':
            # "open" is the straightforward mode, but it locks the file exclusively,
            # so it may fail when the MDD is already open in Professional.
            # mDocument = win32com.client.Dispatch("MDM.Document")
            mDocument = Win32Client.Dispatch("MDM.Document")
            # openConstants_oNOSAVE = 3
            openConstants_oREAD = 1
            # openConstants_oREADWRITE = 2
            print(f'opening MDM document using method "open": "{mdd_path}"')
            # we'll check that the file exists so that the error message is more informative - otherwise you see a long stack of messages that do not tell much
            if not(Path(mdd_path).is_file()):
                raise FileNotFoundError(f'file not found: {mdd_path}')
            mDocument.Open( mdd_path, "", openConstants_oREAD )
            self.__document = mDocument
        elif method=='join':
            # "join" creates a new empty document and merges the opened document into
            # it. It does not need exclusive access, so it is a more stable fallback.
            # mDocument = win32com.client.Dispatch("MDM.Document")
            mDocument = Win32Client.Dispatch("MDM.Document")
            print(f'opening MDM document using method "join": "{mdd_path}"')
            # we'll check that the file exists so that the error message is more informative - otherwise you see a long stack of messages that do not tell much
            if not(Path(mdd_path).is_file()):
                raise FileNotFoundError(f'file not found: {mdd_path}')
            mDocument.Join(mdd_path, "{..}", 1, 32|16|512)
            self.__document = mDocument
        else:
            raise ValueError(f'MDM Open: Unknown open method, {method}')

        self.__mdd_path = mdd_path
        self.__read_datetime = datetime.now()
        config_default = {
            'features': ['label','attributes','properties','translations','scripting'],
            'sections': ['languages','shared_lists','fields','pages','routing'],
            'contexts': ['Question','Analysis']
        }
        self.__config = { **config_default, **config }



    # unlink document if some error happened, or if we are done processing it
    def __del__(self):
        """Close the COM document when the object is garbage collected."""
        if self.__document is not None:
            self.__document.Close()
        print('MDM document closed')



    # strange methods required by python so that I can use "with"
    # I still don't understand why this is needed as we already have __init__ and __del__ and allll should work, why on Earth __enter__ and __exit__ are necessary????
    def __enter__(self):
        """Support ``with MDMDocument(...)`` context management."""
        return self
    def __exit__(self, exc_type, exc_val, exc_tb):
        """No-op; the COM document is released in ``__del__``."""
        pass



    # actually the main method of this class
    # reads the entire MDD and return it as a "report"
    # return type is dict that is suitable to be saved or sent as json
    def read(self,config=None):
        """Read the whole document and return a JSON-ready report dict.

        Returns a dict with ``report_type``, ``source_file``, timestamps,
        ``source_file_metadata``, ``report_scheme`` and a ``sections`` list.

        Args:
            config: Optional overrides. Currently unused here — the effective
                configuration is the one captured in ``__init__``.
        """
        if config is None:
            config = {}

        # prep some variables - list of languages, list of features, columns, etc
        translations_list = [ f'{langcode}' for langcode in self.__document.Languages ]
        self.__translations = translations_list
        flags_list = [ 'data-type:mdd' ]
        columns_list = ['name']
        for feature_spec in self.__config['features']:
            col_add = None
            if feature_spec=='label':
                col_add = 'label'
            elif feature_spec=='attributes':
                col_add = 'attributes'
            elif feature_spec=='properties':
                col_add = 'properties'
            elif feature_spec=='scripting':
                col_add = 'scripting'
            elif feature_spec=='translations':
                col_add = None # TODO: not super beautiful design sorry
                for langcode in translations_list:
                    columns_list.append(f'langcode-{langcode}')
            else:
                raise AttributeError(f'feature type is not recognized: "{feature_spec}"')
            if  col_add:
                columns_list.append(col_add)

        # ok, here's the final result
        # that's what we return
        result = {
            'report_type': 'MDD',
            'source_file': f'{self.__mdd_path}',
            'report_datetime_utc': self.__read_datetime.astimezone(tz=timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
            'report_datetime_local': self.__read_datetime.strftime('%Y-%m-%dT%H:%M:%SZ'),
            'source_file_metadata': [
                { 'name': 'MDD', 'value': f'{self.__mdd_path}' },
                { 'name': 'generated with mdmtoolsap-mddread script version', 'value': f'{config.get("script_version")}', 'role':'sronly' },
            ],
            'report_scheme': {
                'columns': columns_list,
                'column_headers': {},
                'flags': flags_list
            },
            'sections': [],
        }
        for col in result['report_scheme']['columns']:
            if col in std_col_labels:
                result['report_scheme']['column_headers'][col] = std_col_labels[col]
            if re.match(r'^\s*?langcode-',col):
                result['report_scheme']['column_headers'][col] = f"{std_col_labels['translations']}{re.sub(r'^\s*?langcode-','',col)}"
        for section_name in self.__config['sections']:
            section_content = None
            if False:
                pass
            # elif section_name == 'mdmproperties':
            #     section_content = [{'name':'','properties':self.__read_mdm_item_properties(self.__document)}]
            elif section_name == 'languages':
                section_content = self.__read_languages()
            elif section_name == 'shared_lists':
                section_content =  self.__read_sharedlists()
            elif section_name == 'fields':
                zero_item = {
                    'name': '',
                }
                if 'properties' in self.__config['features']:
                    zero_item['properties'] = self.__read_mdm_item_properties(self.__document)
                if 'scripting' in self.__config['features']:
                    zero_item['scripting'] = self.__document.Script
                fields = self.__read_fields(self.__document.Fields)
                section_content = [] +  [zero_item] + fields
            elif section_name == 'pages':
                section_content = self.__read_pages()
            elif section_name == 'routing':
                section_content = self.__read_routing()
            else:
                raise AttributeError(f'unrecognized section type: {section_name}')
            result['sections'].append({'name':section_name,'content':section_content})
        return result

    def __read_languages(self):
        """Return the document languages as rows of ``{name, label}``."""

        try:

            result = []

            config = self.__config
            document = self.__document

            for item in document.Languages:
                try:
                    result_part = {'name': f'{item.Name}'}
                    for read_feature in config['features']:
                        if read_feature=='label':
                            result_part['label'] = f'{item.LongName}'
                    result.append(result_part)
                except Exception as e:
                    try:
                        print(f'{STDOUT_COLOR_RED}Failed when processing language {item.Name}{STDOUT_COLOR_RESET}',file=sys.stderr)
                    except:
                        print(f'{STDOUT_COLOR_RED}Failed when processing language {f"{item}"[:31]}{STDOUT_COLOR_RESET}',file=sys.stderr)
                    raise e

            return result

        except Exception as e:
            print(f'{STDOUT_COLOR_RED}failed when processing languages{STDOUT_COLOR_RESET}',file=sys.stderr)
            raise e

    def __read_sharedlists(self):
        """Return shared-list rows (and their categories/elements when expanded)."""

        try:

            result = []

            # config = self.__config
            document = self.__document
            fields = document.Types

            sharedlists_list = [ f'{slname.Name}' for slname in fields ]
            sharedlists_list.sort()
            for sl_name in sharedlists_list:
                try:
                    item = fields[sl_name]
                    result_item = {
                        **{
                            'name': sl_name,
                            # 'elements': [],
                        },
                        **self.__read_mdm_item(item)
                    }
                    result.append(
                        result_item
                    )
                    result_item = None
                    for cat in self.__list_categories(item):
                        try:
                            #cat_name = '{name}'.format(name=cat.Name)
                            element_item = self.__read_mdm_item(cat)
                            element_item['name'] = f"{sl_name}.categories[{element_item['name']}]"
                            result.append( element_item )
                        except Exception as e:
                            try:
                                print(f'{STDOUT_COLOR_RED}Failed when processing category {cat.Name}{STDOUT_COLOR_RESET}',file=sys.stderr)
                            except:
                                print(f'{STDOUT_COLOR_RED}Failed when processing category {f"{cat}"[:31]}{STDOUT_COLOR_RESET}',file=sys.stderr)
                            raise e
                except Exception as e:
                    try:
                        print(f'{STDOUT_COLOR_RED}Failed when processing shared list {item.Name}{STDOUT_COLOR_RESET}',file=sys.stderr)
                    except:
                        print(f'{STDOUT_COLOR_RED}Failed when processing shared list {f"{sl_name}"[:31]}{STDOUT_COLOR_RESET}',file=sys.stderr)
                    raise e

            return result

        except Exception as e:
            print(f'{STDOUT_COLOR_RED}failed when processing shared lists{STDOUT_COLOR_RESET}',file=sys.stderr)
            raise e

    def __read_pages(self):
        """Return the document pages and their entries as rows."""

        try:

            result = []

            # config = self.__config
            document = self.__document
            fields = document.Pages

            pages_list = [ f'{slname.Name}' for slname in fields ]
            for item_name in pages_list:
                try:
                    item = fields[item_name]
                    result_item = {
                        **{
                            'name': item_name
                        },
                        **self.__read_mdm_item(item)
                    }
                    result.append(
                        result_item
                    )
                    result_item = None
                    for cat in item:
                        try:
                            #cat_name = '{name}'.format(name=cat.Name)
                            item_add = self.__read_mdm_item(cat)
                            item_add['name'] = f"{item_name}.{item_add['name']}"
                            result.append(item_add)
                        except Exception as e:
                            try:
                                print(f'{STDOUT_COLOR_RED}Failed when processing page entry {cat.Name}{STDOUT_COLOR_RESET}',file=sys.stderr)
                            except:
                                print(f'{STDOUT_COLOR_RED}Failed when processing page entry {f"{cat}"[:31]}{STDOUT_COLOR_RESET}',file=sys.stderr)
                            raise e
                except Exception as e:
                    try:
                        print(f'{STDOUT_COLOR_RED}Failed when processing page {item.Name}{STDOUT_COLOR_RESET}',file=sys.stderr)
                    except:
                        print(f'{STDOUT_COLOR_RED}Failed when processing page {f"{item_name}"[:31]}{STDOUT_COLOR_RESET}',file=sys.stderr)
                    raise e

            return result

        except Exception as e:
            print(f'{STDOUT_COLOR_RED}failed when processing pages{STDOUT_COLOR_RESET}',file=sys.stderr)
            raise e

    def __read_fields(self,fields):
        """Flatten the given MDM fields into a list of report rows."""

        try:

            result = []

            # config = self.__config

            fields_list = [ f'{item.Name}' for item in fields ]
            for item_name in fields_list:
                try:
                    item = fields[item_name]
                    result_item = self.__read_process_field(item)
                    #result.append( result_item )
                    result = result + result_item

                except Exception as e:
                    try:
                        print(f'{STDOUT_COLOR_RED}Failed when processing mdm field {item.Name}{STDOUT_COLOR_RESET}',file=sys.stderr)
                    except:
                        print(f'{STDOUT_COLOR_RED}Failed when processing mdm field {f"{item_name}"[:31]}{STDOUT_COLOR_RESET}',file=sys.stderr)
                    raise e

            return result

        except Exception as e:
            print(f'{STDOUT_COLOR_RED}failed when processing fields{STDOUT_COLOR_RESET}',file=sys.stderr)
            raise e

    def __read_process_field(self,item):
        """Process one MDM field and return its rows (self plus nested parts).

        Dispatches on ``ObjectTypeValue`` (question/array/grid/block/...) and emits
        nested categories, subfields and helper fields as additional rows with dotted
        ``name`` prefixes.
        """

        def has_attribute(item,attr):
            return attr.lower() in [m.lower() for m in dir(item)]

        item_name = item.Name
        result_other_items = []
        try:

            result_item = {
                **{
                    'name': f'{item_name}',
                    'attributes': {
                        'object_type_value': item.ObjectTypeValue,
                        #'data_type': item.DataType,
                        #'is_grid': item.IsGrid,
                    },
                },
                **self.__read_mdm_item(item)
            }
            object_type_value = item.ObjectTypeValue
            # ObjectTypeValue constants:
            # 0 = Question
            # 1 = Array
            # 2 = Grid
            # 3 = Class
            # 4 = Element
            # 10 = VariableInstance
            # 16 = Variables
            if object_type_value==0:
                # regular variable
                result_item['attributes']['type'] = 'plain'
                data_type = item.DataType
                result_item['attributes']['data_type'] = data_type
                if data_type==0:
                    # info
                    result_item['attributes']['type'] = 'plain/info'
                elif data_type==1:
                    # long
                    result_item['attributes']['type'] = 'plain/long'
                    result_item['attributes']['minvalue'] = item.MinValue
                    result_item['attributes']['maxvalue'] = item.MaxValue
                elif data_type==2:
                    # text
                    result_item['attributes']['type'] = 'plain/text'
                    result_item['attributes']['minvalue'] = item.MinValue
                    result_item['attributes']['maxvalue'] = item.MaxValue
                elif data_type==3:
                    # categorical
                    result_item['attributes']['type'] = 'plain/categorical'
                    result_item['attributes']['minvalue'] = item.MinValue
                    result_item['attributes']['maxvalue'] = item.MaxValue
                    for cat in self.__list_categories(item):
                        item_add = self.__read_mdm_item(cat)
                        item_add['name'] = f"{item_name}.categories[{item_add['name']}]"
                        result_other_items.append(item_add)
                elif data_type==5:
                    # date
                    result_item['attributes']['type'] = 'plain/date'
                elif data_type==6:
                    # double
                    result_item['attributes']['type'] = 'plain/double'
                    result_item['attributes']['minvalue'] = item.MinValue
                    result_item['attributes']['maxvalue'] = item.MaxValue
                elif data_type==7:
                    # boolean
                    result_item['attributes']['type'] = 'plain/boolean'
                pass
            elif object_type_value==1:
                # array (loop)
                result_item['attributes']['type'] = 'array'
                result_item['attributes']['is_grid'] = item.IsGrid
                for cat in self.__list_categories(item):
                    item_add = self.__read_mdm_item(cat)
                    item_add['name'] = f"{item_name}.categories[{item_add['name']}]"
                    result_other_items.append(item_add)
                for cat in item.Fields:
                    #result_item['attributes']['fields'].append(self.__read_process_field(cat))
                    result_other_items = result_other_items + [ {**item,'name':f"{item_name}.{item['name']}"} for item in self.__read_process_field(cat) ]
            elif object_type_value==2:
                # Grid (it seems it's something different than Array, but I can't understand their logic; maybe it's different because it has a different db setup in case data, I don't know)
                # Execute Error: The '<Object>.IGrid' type does not support the 'categories' property
                result_item['attributes']['type'] = 'grid'
                if has_attribute(item,'IsGrid'):
                    try:
                        # strange, sometimes it crashes here at "IsGrid"
                        # for example, on Mailchimp SMB auto MDD
                        # I'll make it optional - we'll continue execution on AttributeError
                        # it did not crash for me on object_type_value==1 so I am not updating there, but maybe I'l have to
                        result_item['attributes']['is_grid'] = item.IsGrid
                    except AttributeError:
                        pass
                for cat in self.__list_categories(item):
                    item_add = self.__read_mdm_item(cat)
                    item_add['name'] = f"{item_name}.categories[{item_add['name']}]"
                    result_other_items.append(item_add)
                for cat in item.Fields:
                    #result_item['attributes']['fields'].append(self.__read_process_field(cat))
                    result_other_items = result_other_items + [ {**item,'name':f"{item_name}.{item['name']}"} for item in self.__read_process_field(cat) ]
            elif object_type_value==3:
                # class (block)
                result_item['attributes']['type'] = 'block'
                #result_item['attributes']['fields'] = []
                for cat in item.Fields:
                    #result_item['attributes']['fields'].append(self.__read_process_field(cat))
                    result_other_items = result_other_items + [ {**item,'name':f"{item_name}.{item['name']}"} for item in self.__read_process_field(cat) ]
            elif object_type_value==16:
                result_item['attributes']['type'] = 'plain/type16'
                # not sure what is it, an example is Respondent.Serial (in some projects)
                pass
            else:
                raise ValueError(f'unrecognized object data type: {object_type_value}')

            # add couple more attributes necessary for AA and correct data processing
            if has_attribute(item,'IsSystem'):
                if item.IsSystem:
                    result_item['attributes']['is_system'] = 'True'
            if has_attribute(item,'HasCaseData'):
                if item.HasCaseData==False:
                    result_item['attributes']['has_case_data'] = 'False'

            for cat in item.HelperFields:
                #result_item['attributes']['fields'].append(self.__read_process_field(cat))
                result_other_helperfields_items = [ {**item,'name':f"{item_name}.{item['name']}"} for item in self.__read_process_field(cat) ]
                # as a result, we can get a collection of items with all nested elements, categories, inner subfields, etc
                # but the first item is a root item
                # and it should exist
                # and it should receive an attribute that it's a HelperField
                assert len(result_other_helperfields_items)>=1
                result_other_helperfields_items[0]['attributes'].append({'name':'is_helper_field','value':'True'})
                result_other_items = result_other_items + result_other_helperfields_items

            # we need to reformat attributes collection
            attributes_upd = []
            for item_key, item_value in result_item['attributes'].items():
                attributes_upd.append({'name':item_key,'value':f'{item_value}'})
            result_item['attributes'] = attributes_upd

            return [result_item] + result_other_items

        except Exception as e:
            print(f'{STDOUT_COLOR_RED}failed when processing inner field fields from "{item_name}"{STDOUT_COLOR_RESET}',file=sys.stderr)
            raise e

    def __read_routing(self):
        """Return the routing script as a single ``{name, label}`` row."""

        try:

            result = []

            # config = self.__config
            document = self.__document

            for routing_part in ['']:
                val = f'{document.Routing.Script}'
                result.append({'name':routing_part,'label':val})

            return result

        except Exception as e:
            print(f'{STDOUT_COLOR_RED}failed when processing routing{STDOUT_COLOR_RESET}',file=sys.stderr)
            raise e

    def __read_mdm_item_properties(self,item):
        """Return the item's custom properties as a sorted ``[{name, value}, ...]``.

        Properties are read in each configured context and de-duplicated by name; a
        property that differs across contexts is also emitted with a context suffix.
        """

        try:

            document = self.__document
            config = self.__config
            result_properties = []
            context_preserve = document.Contexts.Current
            properties_list = []
            properties = {}
            for read_context in document.Contexts:
                if str(read_context).lower() in [ctx.lower() for ctx in config['contexts']]:
                    document.Contexts.Current = read_context
                    for index_prop in range( 0, item.Properties.Count ):
                        prop_name = f'{item.Properties.Name(index_prop)}'
                        properties_list.append(prop_name)
                        value_sanitized = f'{item.Properties[prop_name]}'
                        # checking for duplicates
                        if not(prop_name in properties):
                            # if not found, all good, we set the value
                            properties[prop_name] = value_sanitized
                        else:
                            # else, if it exists in a different context, we check if it's the same value
                            if properties[prop_name] == value_sanitized:
                                pass # all good
                            else:
                                # if it's a different value (which should never happen)
                                # we add it along with specification of context name
                                # i.e. first value from "Analisis" context will be simply "ShortName"
                                # and we'll add "ShortName (Question)" from "Question" context
                                # it is tested, it is working, but per my understanding this should never happen
                                prop_name_copy = f'{prop_name} ({read_context.Name})'
                                properties_list.append(prop_name_copy)
                                properties[prop_name_copy] = value_sanitized
            document.Contexts.Current = context_preserve
            # remove duplicates
            properties_list = list(set(properties_list))
            # and sort
            properties_list.sort()
            # and return in final format - a list of {name,value} dicts
            for prop_name in properties_list:
                result_properties.append({ 'name': prop_name, 'value': properties[prop_name] })
            return result_properties

        except Exception as e:
            itemname = f'{item}'
            try:
                itemname = item.Name
            except:
                pass
            print(f'{STDOUT_COLOR_RED}failed when making a list of properties of {itemname}{STDOUT_COLOR_RESET}',file=sys.stderr)
            raise e


    def __list_categories(self,item):
        """Yield the item's categories/elements, resolving shared-list references.

        For an element that is a shared-list reference, ``stepinto`` expands the shared
        list into its categories/elements, while ``stepover`` yields only the reference
        (controlled by ``config_sharedlists_listcats_stepinto``).
        """
        config = self.__config
        item_elements = None
        try:
            item_elements = item.Elements
        except AttributeError as e:
            item_name = '<unknown>'
            try:
                item_name = item.Name
            except:
                pass
            raise self.__class__.ErrorItemNotIterable(f"Can't iterate over elements in \"{item_name}\"") from e
        for cat in item_elements:
            try:
                if cat.IsReference:
                    if 'config_sharedlists_listcats_stepinto' in config and config['config_sharedlists_listcats_stepinto']:
                        try:
                            sl_name_clean = re.sub(r'[\^\\/\.]','',cat.ReferenceName,flags=re.I|re.DOTALL)
                            mdmsharedlist = item.Document.Types[sl_name_clean]
                            yield from self.__list_categories(mdmsharedlist)
                        except Exception as e:
                            raise Exception(f'Was not able to refer to a Shared List "{cat.ReferenceName}": {e}') from e
                    else:
                        try:
                            sl_name_clean = re.sub(r'[\^\\/\.]','',cat.ReferenceName,flags=re.I|re.DOTALL)
                            mdmsharedlist = item.Document.Types[sl_name_clean]
                            # yield mdmsharedlist # <- we don't need shared list scripts printed as cat element; we should return cat, not shared list
                            yield cat
                        except Exception as e:
                            raise Exception(f'Was not able to refer to a Shared List "{cat.ReferenceName}": {e}') from e
                elif cat.Type==0:
                    yield cat
                elif (cat.Type==1) or (cat.Type==13):
                    # yield cat
                    yield from self.__list_categories(cat)
                elif (cat.Type==4097) or (re.match(r'^\s*?\d+\s*?$',str(cat.Name),flags=re.I|re.DOTALL)):
                    yield cat
                else:
                    try:
                        yield from self.__list_categories(cat)
                    except self.__class__.ErrorItemNotIterable:
                        yield cat
            except Exception as e:
                try:
                    print(f'{STDOUT_COLOR_RED}Failed when processing element or category: {cat.Name}{STDOUT_COLOR_RESET}',file=sys.stderr)
                except:
                    print(f'{STDOUT_COLOR_RED}Failed when processing element or category: {cat}{STDOUT_COLOR_RESET}',file=sys.stderr)
                raise


    def __read_mdm_item(self,item):
        """Read a single MDM item into a flat dict keyed by the configured features."""

        item_name = f'{item.Name}'

        try:

            result = {
                'name': item_name
            }
            is_shared_list = None
            mdmsharedlist = None
            try:
                is_shared_list = item.IsReference
                ref_name = None
                try:
                    ref_name = f'{item.ReferenceName}'
                except AttributeError:
                    pass
                sl_name_clean = None
                if ref_name:
                    sl_name_clean = re.sub(r'[\^\\/\.]','',ref_name,flags=re.I|re.DOTALL)
                if sl_name_clean:
                    mdmsharedlist = item.Document.Types[sl_name_clean]
                    result['name'] = mdmsharedlist.Name
            except AttributeError:
                pass

            config = self.__config
            #document = self.__document

            for read_feature in config['features']:
                if read_feature=='label':
                    val_label = f'{item.Label}'
                    result[read_feature] = val_label
                elif read_feature=='attributes':
                    pass
                elif read_feature=='translations':
                    # Translation overlays are often trimmed/removed once a project is
                    # in the data-processing/Analysis stage, so translation text is read
                    # through the "Question" context (the primary/safe source).
                #elif read_feature[:9]=='langcode-':
                    #langcode = read_feature[9:]
                    for langcode in self.__translations:
                        #val_label = '{val}'.format(val=item.Labels["Label"].Text["Question"][langcode])
                        val_label = '???'
                        # item.Labels('Question','ENU')
                        try:
                            val_label = f"{item.Labels('Question', langcode)}"
                        except Exception as e:
                            val_label = f'{e}'
                        result[f'langcode-{langcode}'] = val_label
                elif read_feature=='properties':
                    result[read_feature] = self.__read_mdm_item_properties(item)
                elif read_feature=='scripting':
                    #val_label = '{val}'.format(val=item.Script)
                    val_label = '???'
                    try:
                        val_label = f'{item.Script}'
                    except Exception as e:
                        val_label = f'{e}'
                    result[read_feature] = val_label
                else:
                    raise AttributeError(f'unrecognized feature type: {read_feature}')
            return result

        except Exception as e:
            print(f'{STDOUT_COLOR_RED}failed when processing "{item_name}"{STDOUT_COLOR_RESET}',file=sys.stderr)
            raise e



def entry_point(*argcs,**kwargs):
    """CLI entry point: parse args, read the MDD, and write the JSON report next to it."""
    script_name = 'MDD read script'
    try:
        time_start = datetime.now()
        parser = argparse.ArgumentParser(
            description="Read MDD",
            prog='mdmtoolsapram read_mdd'
        )
        parser.add_argument(
            '-1',
            '--mdd',
            help='Input MDD',
            type=str,
            required=True
        )
        parser.add_argument(
            '--method',
            default='open',
            help='"open" (default) is the straightforward mode; "join" creates a new empty document and merges the opened document into it. "open" locks the file exclusively and may fail if it is already in use, while "join" does not require exclusive access (fallback)',
            choices=['open','join'],
            type=str,
            required=False
        )
        parser.add_argument(
            '--config-features',
            help='Comma-separated list of content kinds to process and include in the output JSON (each becomes one or more columns): label, attributes, properties, translations, scripting. Default: label,attributes,properties,translations,scripting',
            type=str,
            required=False
        )
        parser.add_argument(
            '--config-contexts',
            help='Comma-separated list of contexts used when reading labels and custom properties. "Question" usually holds formatted question text (HTML tags); "Analysis" usually holds cleaned/truncated labels for data processing. Default: Question,Analysis',
            type=str,
            required=False
        )
        parser.add_argument(
            '--config-sections',
            help='Comma-separated list of document parts to include in the output JSON: languages, shared_lists, fields, pages, routing. Default: languages,shared_lists,fields,pages,routing',
            type=str,
            required=False
        )
        parser.add_argument(
            '--config-sharedlists-listcats',
            help='How to handle a field that references a shared list instead of listing its own categories/elements: "stepinto" expands the shared list into its full list of categories/elements; "stepover" prints only the reference to the shared list. Both are useful depending on the consumer. Default: stepinto',
            type=str,
            choices = ["stepinto","stepover"],
            required=False
        )
        # args = None
        # args_rest = None
        # if( ('arglist_strict' in config) and (not config['arglist_strict']) ):
        #     args, args_rest = parser.parse_known_args()
        # else:
        args = parser.parse_args(*argcs,**kwargs)
        inp_mdd = None
        if args.mdd:
            inp_mdd = Path(args.mdd)
            inp_mdd = str(inp_mdd.resolve())
        else:
            raise FileNotFoundError('Inp source: file not provided; please use --mdd')

        if not(Path(inp_mdd).is_file()):
            raise FileNotFoundError(f'file not found: {inp_mdd}')

        method = args.method or 'open'

        config = {
            'script_name': script_name,
            'script_started': time_start,
            'script_version': script_version.strip(),
            # 'features': ['label','attributes','properties','translations'], # ,'scripting'],
            'features': ['label','attributes','properties','translations','scripting'],
            'sections': ['languages','shared_lists','fields','pages','routing'],
            'contexts': ['Question','Analysis'],
            'config_sharedlists_listcats_stepinto': CONFIG_LISTCAT_STEPINTO_SL,
        }
        if args.config_features:
            config['features'] = args.config_features.split(',')
        if args.config_contexts:
            config['contexts'] = args.config_contexts.split(',')
        if args.config_sections:
            config['sections'] = args.config_sections.split(',')
        if args.config_sharedlists_listcats:
            config['config_sharedlists_listcats_stepinto'] = (args.config_sharedlists_listcats == 'stepinto')

        print(f'{script_name}: script started at {time_start}')
        print(f'{script_name}: MDD: {inp_mdd}')

        with MDMDocument(inp_mdd,method,config) as doc:

            result = doc.read(config)

            result_json = json.dumps(result, indent=4)
            result_json_fname = ( Path(inp_mdd).parents[0] / f"{Path(inp_mdd).name}.json" if Path(inp_mdd).is_file() else re.sub(r'^\s*?(.*?)\s*?$',lambda m: f"{m[1]}.json",f'{inp_mdd}') )
            print(f'{script_name}: {STDOUT_COLOR_GREEN}saving as "{result_json_fname}"{STDOUT_COLOR_RESET}')
            with open(result_json_fname, "w") as outfile:
                outfile.write(result_json)

        time_finish = datetime.now()
        print(f'{script_name}: finished at {time_finish} (elapsed {time_finish-time_start})')
    except Exception as e:
        # for pretty-printing any issues that happened during runtime; if we hit FileNotFound I don't appreciate when a log traceback is shown, the error should be simple and clear
        # the program is designed to be user-friendly
        # that's why we reformat error messages a little bit
        # stack trace is still printed (I even made it longer to 20 steps!)
        # but the error message itself is separated and printed as the last message again

        # for example, I don't write 'print('File Not Found!');exit(1);', I just write 'raise FileNotFoundErro()'
        print('',file=sys.stderr)
        print('Stack trace:',file=sys.stderr)
        print('',file=sys.stderr)
        traceback.print_exception(e,limit=20)
        print('',file=sys.stderr)
        print('',file=sys.stderr)
        print('',file=sys.stderr)
        print(f'{STDOUT_COLOR_RED}Error:{STDOUT_COLOR_RESET}',file=sys.stderr)
        print('',file=sys.stderr)
        print(f'{STDOUT_COLOR_RED}{e}{STDOUT_COLOR_RESET}',file=sys.stderr)
        print('',file=sys.stderr)
        exit(1)


if __name__ == '__main__':
    entry_point()
