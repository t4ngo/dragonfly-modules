#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for **Firefox**
==============================

This module offers direct control of the `Firefox 
<http://www.mozilla.com/en-US/firefox/>`_ web browser.  It 
requires the `mouseless-browsing 
<https://addons.mozilla.org/en-US/firefox/addon/879>`_ 
(mlb) add-on for reliable access to hyperlinks.

This module includes direct searching using Firefox's 
search bar and Firefox's keyword searching.  It also 
allows single-utterance submitting of text into form text 
fields.

Customization
-------------
Users should customize this module by editing its 
configuration file.  In this file they should edit the 
``search.searchbar`` and ``search.keywords`` settings to 
match their own personal search preferences.  These 
variables map *what you say* to which *search engines* to 
use.

"""


#---------------------------------------------------------------------------

from dragonfly.all import (Grammar, AppContext, Rule, MappingRule,
                           Dictation, Alternative, Choice, RuleRef,
                           Key, Text,
                           Config, Section, Item)
from dragonfly.grammar.integer_en  import Integer, Digits


#---------------------------------------------------------------------------
# Set up this module's configuration.

config                 = Config("Firefox control")
config.search          = Section("Search-related section")
config.search.keywords = Item(
                              default={
                                       "wikipedia": "wikipedia",
                                      },
                              doc="Mapping of spoken-forms to Firefox search-keywords.",
                             )
config.search.searchbar = Item(
                              default=[
                                       "google",
                                       "yahoo",
                                       "amazon",
                                       "answers",
                                       "creative commons",
                                       "eBay",
                                       "wikipedia",
                                      ],
                              doc="Spoken-forms of search engines in the Firefox search-bar; they must be given in the same order here as they are available in Firefox.",
                             )

config.lang                        = Section("Language section")
config.lang.new_win                = Item("new (window | win)")
config.lang.new_tab                = Item("new (tab | sub)")
config.lang.close_tab              = Item("close (tab | sub)")
config.lang.address_bar            = Item("address [bar]")
config.lang.copy_address           = Item("copy address")
config.lang.paste_address          = Item("paste address")
config.lang.search_bar             = Item("search bar")
config.lang.go_home                = Item("go home")
config.lang.stop_loading           = Item("stop loading")
config.lang.toggle_tags            = Item("toggle tags")
config.lang.fresh_tags             = Item("fresh tags")
config.lang.caret_browsing         = Item("(caret | carrot) browsing")
config.lang.bookmark_page          = Item("bookmark [this] page")
config.lang.save_page_as           = Item("save [page | file] as")
config.lang.print_page             = Item("print [page | file]")
config.lang.show_tab_n             = Item("show tab <n>")
config.lang.back                   = Item("back [<n>]")
config.lang.forward                = Item("forward [<n>]")
config.lang.next_tab               = Item("next tab [<n>]")
config.lang.prev_tab               = Item("(previous | preev) tab [<n>]")
config.lang.normal_size            = Item("normal text size")
config.lang.smaller_size           = Item("smaller text size [<n>]")
config.lang.bigger_size            = Item("bigger text size [<n>]")
config.lang.submit                 = Item("submit")
config.lang.submit_text            = Item("submit <text>")
config.lang.find                   = Item("find")
config.lang.find_text              = Item("find <text>")
config.lang.find_next              = Item("find next [<n>]")
config.lang.link_open              = Item("[link] <link> [open]")
config.lang.link_select            = Item("[link] <link> select")
config.lang.link_force             = Item("[link] <link> force")
config.lang.link_window            = Item("[link] <link> window")
config.lang.link_tab               = Item("[link] <link> tab")
config.lang.link_copy              = Item("[link] <link> copy")
config.lang.link_copy_into_tab     = Item("[link] <link> copy into tab")
config.lang.link_list              = Item("[link] <link> list")
config.lang.link_submit            = Item("[link] <link> submit")
config.lang.link_submit_text       = Item("[link] <link> submit <text>")

config.lang.search_text            = Item("[power] search [for] <text>")
config.lang.search_keyword_text    = Item("[power] search <keyword> [for] <text>")
config.lang.search_searchbar_text  = Item("[power] search <searchbar> [for] <text>")

#config.generate_config_file()
config.load()


#---------------------------------------------------------------------------
# Check and prepare search-related config values.

keywords = config.search.keywords
searchbar = dict([(n,i) for i,n in enumerate(config.search.searchbar)])


#---------------------------------------------------------------------------
# Create the rule to match mouseless-browsing link numbers.

class LinkRule(Rule):

    def __init__(self):
        element = Alternative(children=(
            Integer("link_int", 1, 10000),
            Digits("link_dgt", 1, 7),
            ))
        Rule.__init__(self, "link_rule", element, exported=False)

    def value(self, node):
        # Handle link if specified as integer.
        child = node.get_child_by_name("link_int")
        if child: digits = tuple(str(child.value()))
            
        # Handle link if specified as digits.
        child = node.get_child_by_name("link_dgt")
        if child: digits = [str(i) for i in child.value()]
            
        # Format and return keystrokes to select the link.
        link_keys = "f6,s-f6," + ",".join(["numpad"+i for i in digits])
        self._log.debug("Link keys: %r" % link_keys)
        return link_keys

link = RuleRef(name="link", rule=LinkRule())


#---------------------------------------------------------------------------
# Create the main command rule.

class CommandRule(MappingRule):

    mapping = {
        config.lang.new_win:            Key("c-n"),
        config.lang.new_tab:            Key("c-t"),
        config.lang.close_tab:          Key("c-w"),
        config.lang.address_bar:        Key("a-d"),
        config.lang.copy_address:       Key("a-d, c-c"),
        config.lang.paste_address:      Key("a-d, c-v, enter"),
        config.lang.search_bar:         Key("c-k"),
        config.lang.go_home:            Key("a-home"),
        config.lang.stop_loading:       Key("escape"),
        config.lang.toggle_tags:        Key("f12"),
        config.lang.fresh_tags:         Key("f12, f12"),
        config.lang.caret_browsing:     Key("f7"),
        config.lang.bookmark_page:      Key("c-d"),
        config.lang.save_page_as:       Key("c-s"),
        config.lang.print_page:         Key("c-p"),

        config.lang.show_tab_n:         Key("0, %(n)d, enter"),
        config.lang.back:               Key("a-left/15:%(n)d"),
        config.lang.forward:            Key("a-right/15:%(n)d"),
        config.lang.next_tab:           Key("c-tab:%(n)d"),
        config.lang.prev_tab:           Key("cs-tab:%(n)d"),

        config.lang.normal_size:        Key("a-v/20, z/20, r"),
        config.lang.smaller_size:       Key("c-minus:%(n)d"),
        config.lang.bigger_size:        Key("cs-equals:%(n)d"),

        config.lang.submit:             Key("enter"),
        config.lang.submit_text:        Text("%(text)s") + Key("enter"),

        config.lang.find:               Key("c-f"),
        config.lang.find_text:          Key("c-f") + Text("%(text)s"),
        config.lang.find_next:          Key("f3/10:%(n)d"),

        config.lang.link_open:          Key("%(link)s, enter"),
        config.lang.link_select:        Key("%(link)s, shift"),
        config.lang.link_force:         Key("%(link)s, shift/10, enter"),
        config.lang.link_window:        Key("%(link)s, shift/10, apps/20, w"),
        config.lang.link_tab:           Key("%(link)s, shift/10, apps/20, t"),
        config.lang.link_copy:          Key("%(link)s, shift/10, apps/20, a"),
        config.lang.link_copy_into_tab: Key("%(link)s, shift/10, apps/20, a/10, c-t/20, c-v, enter"),
        config.lang.link_list:          Key("%(link)s, enter, a-down"),
        config.lang.link_submit:        Key("%(link)s, enter/30, enter"),
        config.lang.link_submit_text:   Key("%(link)s, enter/30")
                                         + Text("%(text)s") + Key("enter"),

        config.lang.search_text:        Key("c-k")
                                         + Text("%(text)s") + Key("enter"),
        config.lang.search_searchbar_text: Key("c-k, c-up:20, c-down:%(searchbar)d")
                                         + Text("%(text)s") + Key("enter"),
        config.lang.search_keyword_text: Key("a-d")
                                         + Text("%(keyword)s %(text)s")
                                         + Key("enter"),
        }
    extras = [
        link,
        Integer("n", 1, 20),
        Dictation("text"),
        Choice("keyword", keywords),
        Choice("searchbar", searchbar),
        ]
    defaults = {
        "n": 1,
        }


#---------------------------------------------------------------------------
# Create and load this module's grammar.

context = AppContext(executable="firefox")
grammar = Grammar("firefox_general", context=context)
grammar.add_rule(CommandRule())
grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
