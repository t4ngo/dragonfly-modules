#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for **variable name formatting**
===============================================

This module offers easy formatting of variable names.  It 
currently supports the following naming styles:

 - **"score"**: ``big_long_name``
 - **"studley"**: ``BigLongName``
 - **"[all] one word"**: ``biglongname``
 - **"java method"**: ``bigLongName``

Command
-------

Command: **"<format> <dictation>"**
    Format the dictated words with the given format.


Examples
--------

 - Say **"score long variable name"** -> *long_variable_name*
 - Say **"Java method string buffer"** -> *stringBuffer*

"""


#---------------------------------------------------------------------------

from dragonfly.all import (Grammar, CompoundRule,
                           Dictation, Choice, Text)


#---------------------------------------------------------------------------

class HandlerBase(object):
    spec = None
    def handle_text(self, text):
        pass


class UnderscoreHandler(HandlerBase):
    spec = "score"
    def handle_text(self, text):
        output = "_".join(text.split(" "))
        Text(output).execute()

class CFunctionHandler(HandlerBase):
    spec = "under func"
    def handle_text(self, text):
        output = "_".join(text.split(" ")) + "()"
        Text(output).execute()

class StudleyHandler(HandlerBase):
    spec = "studley"
    def handle_text(self, text):
        words = [word.capitalize() for word in text.split(" ")]
        output = "".join(words)
        Text(output).execute()

class OneWordHandler(HandlerBase):
    spec = "[all] one word"
    def handle_text(self, text):
        output = "".join(text.split(" "))
        Text(output).execute()

class UpperOneWordHandler(HandlerBase):
    spec = "one word upper"
    def handle_text(self, text):
        words = [word.upper() for word in text.split(" ")]
        output = "".join(words)
        Text(output).execute()

class UnderscoreUpperHandler(HandlerBase):
    spec = "upper score"
    def handle_text(self, text):
        words = [word.upper() for word in text.split(" ")]
        output = "_".join(words)
        Text(output).execute()

class JavaMethodHandler(HandlerBase):
    spec = "java method"
    def handle_text(self, text):
        words = text.split(" ")
        output = words[0] + "".join(w.capitalize() for w in words[1:])
        Text(output).execute()


#---------------------------------------------------------------------------
# Create the main command rule.

class CommandRule(CompoundRule):

    handlers = {}
    for value in globals().values():
        try:
            if issubclass(value, HandlerBase) and value is not HandlerBase:
                instance = value()
                handlers[instance.spec] = instance.handle_text
        except TypeError:
            continue

    spec     = "<handler> <text>"
    extras   = [
                Choice("handler", handlers),
                Dictation("text"),
               ]

    def _process_recognition(self, node, extras):
        handler = extras["handler"]
        text = extras["text"]
        handler(text)


#---------------------------------------------------------------------------

grammar = Grammar("Variable format")
grammar.add_rule(CommandRule())
grammar.load()

def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
