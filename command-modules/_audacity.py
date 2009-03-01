#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Biddy
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for **Audacity**, an audio-editing application
=============================================================
This module offers voice-commands to control 
`Audacity <http://audacity.sourceforge.net/>`_, an open 
source application for recording and editing sounds.

Customization
-------------
Users can edit the spoken-form of this module's commands
in its configuration file.  This is useful for
translations, for example.

"""

import pkg_resources
pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r76")

from dragonfly import (Grammar, AppContext, MappingRule, Key,
                       Config, Section, Item)


#---------------------------------------------------------------------------
# Initialize this module's configuration.

config = Config("Audacity control")

config.lang                    = Section("Language section")
config.lang.new_project        = Item("[start] new project", doc="Spec starts a new project.")
config.lang.open_project       = Item("open project", doc="Spec opens an existing project.")
config.lang.close_project      = Item("close [this] project", doc="Spec closes currently open project.")
config.lang.save_project       = Item("save [this] project", doc="Spec saves currently open project.")
config.lang.selection_tool     = Item("[open] selection tool", doc="Spec opens the selection tool.")
config.lang.envelope_tool      = Item("[open] envelope tool", doc="Spec opens the envelope tool.")
config.lang.edit_tool          = Item("[open] edit tool", doc="Spec opens the edit tool.")
config.lang.zoom_tool          = Item("[open] zoom tool", doc="Spec opens the zoom tool.")
config.lang.timeshift_tool     = Item("[open] timeshift tool", doc="Spec opens the timeshift tool.")
config.lang.multi_tool         = Item("[open] multi tool", doc="Spec opens the multitool.")
config.lang.next_toolbar_tool  = Item("[move to | go to] next tool", doc="Spec move to the next tool in toolbar.")
config.lang.previous_toolbar_tool = Item("[move to | go to] previous tool", doc="Spec move to the previous tool in toolbar.")
config.lang.silence            = Item("[insert | create] silence", doc="Spec insert silence at current position.")
config.lang.duplicate          = Item("[make] duplicate", doc="Spec makes duplicate.")
config.lang.find_zero_crossings= Item("[find | search] zero crossings", doc="Spec finds zero crossings.")
config.lang.play_stop          = Item("[start | stop] playing", doc="Spec starts or stops playback.")
config.lang.loop               = Item("[start] loop", doc="Spec starts a loop.")
config.lang.pause              = Item("pause [playback | recording]", doc="Spec pauses current playback or recording.")
config.lang.record             = Item("[start | begin] recording", doc="Spec starts recording.")
config.lang.play_cursor_selection = Item("[play] cursor selection", doc="Spec plays one second at cursor position.")
config.lang.zoom_in            = Item("zoom in", doc="Spec zooms in on current view.")
config.lang.zoom_normal        = Item("[return to] standard zoom level", doc="Spec zoom in or out to standard zoom level.")
config.lang.zoom_out           = Item("zoom out", doc="Spec zooms out on current view.")
config.lang.fit_window         = Item("fit [track] to window", doc="Spec fits the track in current window size.")
config.lang.fit_vertically     = Item("fit all tracks [to window]", doc="Spec makes all tracks visible in current window size.")
config.lang.zoom_selection     = Item("zoom [in] selection", doc="Spec zooms in on your current selection.")
config.lang.import_audio       = Item("import audio", doc="Spec imports audio from your PC.")
config.lang.create_label       = Item("[create | make] label", doc="Spec creates a label.")
config.lang.repeat_last_effect = Item("repeat last effect", doc="Spec repeats the last effect.")
config.lang.ripple_delete      = Item("ripple delete selection", doc="Spec deletes current selection and removes empty space.")
config.lang.split_delete       = Item("split delete selection", doc="Spec deletes current selection while leaving space empty.")
config.lang.split_new          = Item("split new", doc="Spec splits new.")
config.lang.join               = Item("join track", doc="Spec removes all empty spaces from track.")
config.lang.disjoin            = Item("disjoin track", doc="Spec brings back empty spaces in track.")
config.lang.mute_all_tracks    = Item("mute all [tracks]", doc="Spec mutes all tracks in file.")
config.lang.unmute_all_tracks  = Item("unmute all [tracks]", doc="Spec unmutes all tracks in file.")
config.lang.collapse_all_tracks= Item("collapse all [tracks]", doc="Spec collapses all tracks in file.")
config.lang.expand_all_tracks  = Item("expand all [tracks]", doc="Spec expands all tracks in file.")
config.lang.export_track       = Item("export track", doc="Spec exports track to preferred file format.")
config.lang.export_selection   = Item("export selection", doc="Spec exports selection to preferred file format.")

#config.generate_config_file()
config.load()


#---------------------------------------------------------------------------
# Create this module's grammar and the context under which it'll be active.

context = AppContext(executable="audacity")
grammar = Grammar("audacity", context=context)


#---------------------------------------------------------------------------
# This module's main keystroke mapping rule.

keystrokes_rule = MappingRule(
    name="keystrokes",
    mapping={
             config.lang.new_project:            Key("c-n"),
             config.lang.open_project:           Key("c-o"),
             config.lang.close_project:          Key("c-w"),
             config.lang.save_project:           Key("c-s"),
             config.lang.selection_tool:         Key("f1"),
             config.lang.envelope_tool:          Key("f2"),
             config.lang.edit_tool:              Key("f3"),
             config.lang.zoom_tool:              Key("f4"),
             config.lang.timeshift_tool:         Key("f5"),
             config.lang.multi_tool:             Key("f6"),
             config.lang.next_toolbar_tool:      Key("d"),
             config.lang.previous_toolbar_tool:  Key("a"),
             config.lang.silence:                Key("c-l"),
             config.lang.duplicate:              Key("c-d"),
             config.lang.find_zero_crossings:    Key("z"),
             config.lang.play_stop:              Key("space"),
             config.lang.loop:                   Key("l"),
             config.lang.pause:                  Key("p"),
             config.lang.record:                 Key("r"),
             config.lang.play_cursor_selection:  Key("b"),
             config.lang.zoom_in:                Key("c-1"),
             config.lang.zoom_normal:            Key("c-2"),
             config.lang.zoom_out:               Key("c-3"),
             config.lang.fit_window:             Key("c-f"),
             config.lang.fit_vertically:         Key("cs-f"),
             config.lang.zoom_selection:         Key("c-e"),
             config.lang.import_audio:           Key("cs-i"),
             config.lang.create_label:           Key("c-b"),
             config.lang.repeat_last_effect:     Key("c-r"),
             config.lang.ripple_delete:          Key("c-k"),
             config.lang.split_delete:           Key("ca-k"),
             config.lang.split_new:              Key("ca-i"),
             config.lang.join:                   Key("c-j"),
             config.lang.disjoin:                Key("ca-j"),
             config.lang.mute_all_tracks:        Key("c-u"),
             config.lang.unmute_all_tracks:      Key("cs-u"),
             config.lang.collapse_all_tracks:    Key("cs-c"),
             config.lang.expand_all_tracks:      Key("cs-x"),

             # Customised hotkeys.
             # (editable in Audacity; Preferences; Keyboard)
             config.lang.export_track:           Key("ca-e"),
             config.lang.export_selection:       Key("cs-e"),
            },
    )

# Add the action rule to the grammar instance.
grammar.add_rule(keystrokes_rule)



#---------------------------------------------------------------------------
# Load the grammar instance and define how to unload it.

grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
