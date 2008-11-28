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

import dragonfly.engines.engine
engine = dragonfly.engines.engine.get_engine()
from dragonfly.all import (Grammar, CompoundRule,
                           Dictation, Choice, Key, Text)


#---------------------------------------------------------------------------

(normal, no_space) = range(2)

class HandlerBase(object):

    spec = None
    spacing = no_space

    def handle_text(self, text):
        if self.spacing == normal:
            engine.mimic(["test"])
            Key("backspace:4").execute()
        formatted = self.format_text(text)
        Text(formatted).execute()

    def format_text(self, text):
        pass


class UnderscoreHandler(HandlerBase):
    spec = "score"
    def format_text(self, text):
        return "_".join(text.split(" "))

class CFunctionHandler(HandlerBase):
    spec = "under func"
    def format_text(self, text):
        return "_".join(text.split(" ")) + "()"

class StudleyHandler(HandlerBase):
    spec = "studley"
    def format_text(self, text):
        words = [word.capitalize() for word in text.split(" ")]
        return "".join(words)

class OneWordHandler(HandlerBase):
    spec = "[all] one word"
    def format_text(self, text):
        return "".join(text.split(" "))

class UpperOneWordHandler(HandlerBase):
    spec = "one word upper"
    def format_text(self, text):
        words = [word.upper() for word in text.split(" ")]
        return "".join(words)

class UnderscoreUpperHandler(HandlerBase):
    spec = "upper score"
    def format_text(self, text):
        words = [word.upper() for word in text.split(" ")]
        return "_".join(words)

class JavaMethodHandler(HandlerBase):
    spec = "Java method"
    def format_text(self, text):
        words = text.split(" ")
        return words[0] + "".join(w.capitalize() for w in words[1:])


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
