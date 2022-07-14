import json

from argparse import ArgumentParser, Namespace
from typing import Dict, List
from winreg import (
  ConnectRegistry,
  EnumKey,
  EnumValue,
  HKEY_LOCAL_MACHINE,
  KEY_READ,
  OpenKey,
  HKEYType,
  QueryInfoKey,
)


def fuzzy_match(a: str, b:str) -> bool:
    """Return true when the stripped, lowercase value of a contains the stripped
    lowercase value of b or vice-versa.
    """
    a_ = a.lower().strip()
    b_ = b.lower().strip()
    return a_ in b_ or b_ in a_


def find_application_display_name_by_name(name_guess: str) -> List[Dict[str, str]]:
    """Returns a list of dictionary values of every item in the two windows registry
    hives where uninstall strings live where the registry key's DisplayName value
    successfully returns a fuzzy match against 'name_guess'

    Args:
        name_guess (str): Guess at an application's DisplayName

    Returns:
        List[Dict[str, str]]: Registry key info for each match
    """
    uninstall_hives = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    matches = []
    with ConnectRegistry(None, HKEY_LOCAL_MACHINE) as hive:
        for sub_key in uninstall_hives[:1]:
            matches += search_hive(name_guess, hive, sub_key)

    return matches


def search_hive(name_guess: str, hive: HKEYType, sub_key: HKEYType) -> List[Dict[str, str]]:
    """Given a hive to iterate, returns a List[Dict] of registry keys where the 
    DisplayName key fuzzy matches 'name guess'.

    Args:
        name_guess (str): Guess at an application's DisplayName
        hive (HKEYType): the uninstallation hive
        sub_key (HKEYType): 

    Returns:
        List[Dict[str, str]]: matches 
    """
    matches = []
    with OpenKey(hive, sub_key, 0, KEY_READ) as uninstall_key:
        application_keys = []
        num_of_keys = QueryInfoKey(uninstall_key)[0]
        for i in range(num_of_keys):
            application_keys.append(EnumKey(uninstall_key, i))
            
        for application_key in application_keys:
            with OpenKey(uninstall_key, application_key, 0, KEY_READ) as application_info_key:
                num_of_values = QueryInfoKey(application_info_key)[1]
                for i in range(num_of_values):
                    values = EnumValue(application_info_key, i)
                    k, v = values[:-1]
                        
                    if k.lower().strip() == "displayname" and fuzzy_match(name_guess, v):
                        matches.append(get_key_details(application_info_key))
        return matches


def get_key_details(application_info_key: HKEYType) -> Dict[str, str]:
    """Return a Dict of registry keys and values from a registry hive

    Args:
        application_info_key (_type_): the registry key to draw values from

    Returns:
        Dict[str, str]: Registry key and value
    """
    details = {}
    num_of_values = QueryInfoKey(application_info_key)[1]
    for i in range(num_of_values):
        values = EnumValue(application_info_key, i)
        k, v = values[:-1]
        details[k] = v
    return details


def get_args() -> Namespace:
    """Get arguments from script input - could just as easily not be wrapped in it's own function.

    Returns:
        Namespace: the args
    """
    parser = ArgumentParser()
    parser.add_argument("guess", type=str, help="Guess at application display name.")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()
    return args


def main() -> None:
    """The entrypoint to the script."""
    args = get_args()
    guess = args.guess
    verbose = args.verbose
    matches = find_application_display_name_by_name(guess)
    for m in matches:
        if verbose:
            print(json.dumps(m, indent=2, sort_keys=True))
        else:
            print(m["DisplayName"])


if __name__ == "__main__":
    main()