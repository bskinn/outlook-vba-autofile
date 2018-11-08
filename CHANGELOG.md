# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2018-11-08

Existing production version released as open source.

### Current Features

 * Numerous elements of the tool are configurable at invocation time,
   allowing setup of different Ribbon buttons for multiple filing
   types/destinations
 * Sets of target folders can be at an arbitrary level of folder nesting
   below the root folder, configurable at invocation time
 * The filing operation can be set up to be invoked from either an Explorer
   or an Inspector; invocation from an Explorer will perform the move on
   all selected items
 * All moved items are (1) marked as read and (2) labeled with a color category;
   the category applied can either be a set category for all individual destinations,
   or a custom category for each destination as defined by a regular expression


