'''
Created on Mar 11, 2020

@author: Gupta Jitendra
'''
from configparser import ConfigParser, NoSectionError
import os
from util import utility_method_class

filepath = os.path.join(utility_method_class.projectDir(),
                        'config\\config.ini')  # Join two separate path in Python but the path name should not we start with / slash

MS_Patch_filepath = os.path.join(utility_method_class.projectDir(), 'config\\mspatch.ini')


def readPropertyKeyValue(section, key):
    cfg = ConfigParser()
    cfg.read(filepath)
    # utility_method_class.customLog().info('File path of the configuration---->   '+ filepath)
    try:
        if section in cfg.sections():
            value = cfg.get(section, key)
            return value

    except FileNotFoundError as e:
        utility_method_class.customLog().error('File not found' + e)
    except NoSectionError as e:
        utility_method_class.customLog().exception('Defined section is not present ' + e)

    else:
        utility_method_class.customLog().info('Unable to process the .ini config file ')


def mspatchReadPropertyKeyValue(section, key):
    cfg = ConfigParser()
    cfg.read(MS_Patch_filepath)
    # utility_method_class.customLog().info('File path of the configuration---->   '+ filepath)
    try:
        if section in cfg.sections():
            value = cfg.get(section, key)
            return value

    except FileNotFoundError as e:
        utility_method_class.customLog().error('File not found' + e)
    except NoSectionError as e:
        utility_method_class.customLog().exception('Defined section is not present ' + e)

    else:
        utility_method_class.customLog().info('Unable to process the .ini config file ')
