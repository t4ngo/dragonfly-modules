#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for **Microsoft Outlook** control
================================================

This module offers a number of commands for direct and efficient
control of Microsoft outlook.

Commands
--------

Command: **"(folder | show me) <folder>"**
    Jump straight to the specified folder.

Command: **"move to <folder>"**
    Move the selected item(s) to the specified folder.

Command: **"save attachments"**
    Save all attachments of selected items.

Command: **"open attachment <n>"**
    Open attachment number <n> within the selected item.

Command: **"[create] new <type>"**
    Creates a new item of the specified type.

Command: **"(synchronize | update) (folders | folder list)"**
    Update this module's internal list of Outlook folders.

Command: **"new (email | mail) [to <addresses>]"**
    Create a new mail item.

Command: **"forward [email | mail] [to <addresses>]"**
    Forward to a selected mail item.

Commands for responding to a meeting request:
     - **"accept and send"**
     - **"accept and edit"**
     - **"accept without response"**
     - **"decline and send"**
     - **"decline and edit"**
     - **"decline without response"**
     - **"tentative and send"**
     - **"tentative and edit"**
     - **"tentative without response"**
     - **"check calendar"**

"""

import pkg_resources
pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r76")

import tempfile
import os
import os.path
import subprocess
from win32com.client  import constants
from pywintypes       import com_error

from dragonfly import (ConnectionGrammar, AppContext, DictListRef, Choice,
                       CompoundRule, MappingRule, DictList, RuleRef,
                       Config, Section, Item, Integer, Key, Text, Pause)


#---------------------------------------------------------------------------
# Set up this module's configuration.

config = Config("Outlook control")
config.lang                 = Section("Language section")
config.lang.go_to_folder    = Item("(folder | show me) <folder>")
config.lang.move_to_folder  = Item("move to <folder>")
config.lang.save_attachments = Item("save (attachments | edges)")
config.lang.open_attachment = Item("open (attachment | edge) <n>")
config.lang.create_new_item = Item("[create] new <type>")
config.lang.sync_folders    = Item("(synchronize | update) (folders | folder list)")
config.lang.item_type_mail  = Item("(mail | email)")
config.lang.item_type_task  = Item("task")
config.lang.new_email       = Item("new (email | mail) [to <addresses>]")
config.lang.forward_email   = Item("forward [email | mail] [to <addresses>]")
config.lang.address_and_word = Item("[and]")
config.lang.meeting_request_actions = Item({
      "accept and send":             Key("a-a/10, c/20, s/10, enter"),
      "accept and edit":             Key("a-a/10, c/20, e/10, enter"),
      "accept without response":     Key("a-a/10, c/20, d/10, enter"),
      "decline and send":            Key("a-a/10, e/20, s/10, enter"),
      "decline and edit":            Key("a-a/10, e/20, e/10, enter"),
      "decline without response":    Key("a-a/10, e/20, d/10, enter"),
      "tentative and send":          Key("a-a/10, a/20, s/10, enter"),
      "tentative and edit":          Key("a-a/10, a/20, e/10, enter"),
      "tentative without response":  Key("a-a/10, a/20, d/10, enter"),
      "check calendar":              Key("a-a/10, h"),
     }, namespace={"Key": Key})
config.contacts             = Section("Contacts section")
config.contacts.addresses   = Item({
      "someone": "someone@example.com",
     })
#config.generate_config_file()
config.load()


#---------------------------------------------------------------------------
# Utility generator function for iterating over COM collections.

def collection_iter(collection):
    for index in xrange(1, collection.Count + 1):
        yield collection.Item(index)

#---------------------------------------------------------------------------
# This module's main grammar.

class ControlGrammar(ConnectionGrammar):

    def __init__(self):
        self.folders = DictList("folders")
        ConnectionGrammar.__init__(
            self,
            name="outlook control",
            context=AppContext(executable="outlook"),
            app_name="Outlook.Application"
           )

    def connection_up(self):
        # Made connection with Outlook -> retrieves available folders.
        self.update_folders()

    def connection_down(self):
        # Lost connection with Outlook -> empty folders list.
        self.reset_folders()

    def update_folders(self):
        namespace = self.application.GetNamespace("MAPI")
        inbox_folder = namespace.GetDefaultFolder(constants.olFolderInbox)
        root_folder = inbox_folder.Parent
        self.folders.set({})
        stack = [collection_iter(root_folder.Folders)]
        while stack:
            try:
                folder = stack[-1].next()
            except StopIteration:
                stack.pop()
                continue
            self.folders[folder.Name] = folder
            stack.append(collection_iter(folder.Folders))

    def reset_folders(self):
        self.folders.set({})

    def get_active_explorer(self):
        try:
            explorer = self.application.ActiveExplorer()
        except com_error, e:
            self._log.warning("%s: COM error getting active explorer: %s"
                              % (self, e))
            return None

        if not explorer:
            self._log.warning("%s: no active explorer." % self)
            return None
        return explorer

grammar = ControlGrammar()


#---------------------------------------------------------------------------
# Synchronize folders rule for explicitly updating folder list.

class SynchronizeFoldersRule(CompoundRule):
    spec = config.lang.sync_folders
    def _process_recognition(self, node, extras):
        self.grammar.update_folders()

grammar.add_rule(SynchronizeFoldersRule())


#---------------------------------------------------------------------------

class GoToFolderRule(CompoundRule):

    spec = config.lang.go_to_folder
    extras = [DictListRef("folder", grammar.folders)]

    def _process_recognition(self, node, extras):
        folder = extras["folder"]

        # Get the currently active explorer.
        explorer = self.grammar.get_active_explorer()
        if not explorer: return

        # Make the explorer view the given folder.
        explorer.SelectFolder(folder)

grammar.add_rule(GoToFolderRule())


#---------------------------------------------------------------------------

class MoveToFolderRule(CompoundRule):

    spec = config.lang.move_to_folder
    extras = [DictListRef("folder", grammar.folders)]

    def _process_recognition(self, node, extras):
        folder = extras["folder"]

        # Get the currently active explorer.
        explorer = self.grammar.get_active_explorer()
        if not explorer: return

        # Move the selected items to the given folder.
        for item in collection_iter(explorer.Selection):
            self._log.debug("%s: moving item %r to folder %r."
                            % (self, item.Subject, folder.Name))
            item.Move(folder)

grammar.add_rule(MoveToFolderRule())


#---------------------------------------------------------------------------

class SaveAttachmentsRule(CompoundRule):

    spec = config.lang.save_attachments

    def _process_recognition(self, node, extras):

        # Get the currently active explorer.
        explorer = self.grammar.get_active_explorer()
        if not explorer: return

        # Save the attachments of the selected items.
        temp_dir = tempfile.mkdtemp()
        print "temporary directory:", temp_dir
        for item in collection_iter(explorer.Selection):
            self._log.debug("%s: saving attachments of item %r."
                            % (self, item.Subject))

            for attachment in collection_iter(item.Attachments):
                print "attachment %r" % attachment.FileName
                filename = os.path.basename(attachment.FileName)
                path = os.path.join(temp_dir, filename)
                attachment.SaveAsFile(path)

        # Open a file browser to the containing directory.
        subprocess.Popen(["explorer", temp_dir])

grammar.add_rule(SaveAttachmentsRule())


#---------------------------------------------------------------------------

class OpenAttachmentRule(CompoundRule):

    spec = config.lang.open_attachment
    extras = [Integer("n", 1, 10)]

    def _process_recognition(self, node, extras):
        index = extras["n"]

        # Get the currently active explorer.
        explorer = self.grammar.get_active_explorer()
        if not explorer: return

        # Make sure that exactly 1 item is selected.
        if explorer.Selection.Count < 0:
            self._log.warning("%s: no selected, not opening." % self)
            return
        elif explorer.Selection.Count > 1:
            self._log.warning("%s: multiple items selected, not opening."
                              % self)
            return

        # Retrieve the attachment of the selected item.
        item = explorer.Selection.Item(1)
        if not (1 <= index <= item.Attachments.Count):
            self._log.warning("%s: attachment index %d of item %r"
                              " out of range (1 <= index <= %d)."
                        % (self, index, item.Subject,
                           item.Attachments.Count))
            return

        attachment = item.Attachments.Item(index)
        self._log.debug("%s: opening attachment %r of item %r."
                        % (self, attachment.FileName, item.Subject))
        filename = os.path.basename(attachment.FileName)

        # Save the attachment to a temporary directory.
        temp_dir = tempfile.mkdtemp()
        path = os.path.join(temp_dir, filename)
        attachment.SaveAsFile(path)

        # Open a file browser to the containing directory.
        os.startfile(path)

grammar.add_rule(OpenAttachmentRule())


#---------------------------------------------------------------------------

class MeetingRequestRule(MappingRule):

    mapping = config.lang.meeting_request_actions

    def _process_recognition(self, value, extras):

        # Get the currently active explorer.
        explorer = self.grammar.get_active_explorer()
        if not explorer: return

        # Save the attachments of the selected items.
        if explorer.Selection.Count != 1:
            self._log.warning("%s: cannot accept meeting requests when"
                              " multiple items are selected.")
            return

        # Retrieve the item.
        item = explorer.Selection.Item(1)

        # Perform spoken action on this meeting request.
        value.execute()

grammar.add_rule(MeetingRequestRule())


#---------------------------------------------------------------------------
# Simple new item rule.

class NewItemRule(CompoundRule):

    spec = config.lang.create_new_item
    types = {        # Use delayed constant lookups, because these
                     #  are loaded when the com connection is set up.
             config.lang.item_type_task:    "olTaskItem",
            }
    extras = [Choice("type", types)]

    def _process_recognition(self, node, extras):
        item_type = getattr(constants, extras["type"])
        item = self.grammar.application.CreateItem(item_type)
        item.Display()

grammar.add_rule(NewItemRule())


#---------------------------------------------------------------------------
# Address list rule and element.

class AddressListRule(MappingRule):

    exported = False
    mapping = config.contacts.addresses


class Addresses(CompoundRule):

    exported = False
    max_names = 3
    spec = "<name>" \
           + (" [" + config.lang.address_and_word + " <name>") * (max_names-1) \
           + "]" * (max_names-1)
    extras = [RuleRef(name="name", rule=AddressListRule())]

    def value(self, node):
        children = node.get_children_by_name("name")
        return [child.value() for child in children]

addresses = RuleRef(name="addresses", rule=Addresses())


#---------------------------------------------------------------------------
# Voice command for creating new (addressed) e-mails.

class NewMailRule(CompoundRule):

    spec = config.lang.new_email
    extras = [addresses]

    def _process_recognition(self, node, extras):
        item_type = getattr(constants, "olMailItem")
        item = self.grammar.application.CreateItem(item_type)
        item.Display()

        if "addresses" in extras:
            addresses = extras["addresses"]
            action = Pause("100") + Text("; ".join(addresses) + ";") + Key("tab:2")
            action.execute()

grammar.add_rule(NewMailRule())


#---------------------------------------------------------------------------
# Voice command for forwarding selected mail items.

class ForwardRule(CompoundRule):

    spec = config.lang.forward_email
    extras = [addresses]

    def _process_recognition(self, node, extras):
        # Get the currently active explorer.
        explorer = self.grammar.get_active_explorer()
        if not explorer:
            return

        # Forward to first item of the current selection.
        for item in collection_iter(explorer.Selection):
            # Create a forwarded copy and display it.
            message = item.Forward()
            message.Display()

            # Insert addresses if given.
            if "addresses" in extras:
                addresses = extras["addresses"]
                action = Pause("100") + Text("; ".join(addresses) + ";") + Key("tab:3")
                action.execute()

            # Only handle the first selected item.
            break

grammar.add_rule(ForwardRule())


#---------------------------------------------------------------------------

class RetrieveContactsRule(CompoundRule):

    spec = "retrieve Outlook contacts"

    def _process_recognition(self, node, extras):
        namespace = self.grammar.application.GetNamespace("MAPI")
        addresses = namespace.AddressLists
        entries = []
        for address_list in collection_iter(addresses):
            print "address list:", address_list.Name
            for entry in collection_iter(address_list.AddressEntries):
                print "    entry name: %r" % entry.Name
                entries.append(entry)
        names = [e.Name for e in entries]
        lines = ["    %(name)20r: %(name)r," % {"name": n} for n in names]
        lines.insert(0, "contacts.names = {")
        lines.append("    }")
        output = "\n".join(lines)
        print output

grammar.add_rule(RetrieveContactsRule())


#---------------------------------------------------------------------------
# Load the grammar instance and define how to unload it.

grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
