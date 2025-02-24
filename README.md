# Wordo Module

## Overview
Wordo is a module component of Spector, an LLM-based computer control program. This specific module handles Microsoft Word automation and control through natural language commands.

## Description
This module provides programmatic control over Microsoft Word applications, allowing for:
- Document manipulation
- Text formatting
- Content insertion
- UI control
- Document state monitoring
- Accessibility information gathering

## Dependencies
```python
import win32com.client
import win32gui
import win32api
import win32con
import json
import re
```

## Key Components

### CommandParser
- Parses natural language commands into executable actions
- Handles command validation and parameter extraction

### WordController
- Core control interface for Microsoft Word
- Manages document operations and text manipulation
- Handles formatting and content insertion

### WordScraper & WordEnhancedScraper
- UI element detection and state monitoring
- Accessibility information gathering
- Document and application state analysis

## Usage Example
```python
from wordo import Wordo

wordo = Wordo()
wordo.run()

# Example commands:
# > start_word
# > create_new
# > write text="Hello World"
# > bold value=True
# > save path="document.docx"
```

## Integration
This module is designed to be integrated with the larger Spector system. It should not be used as a standalone application.

## Note
This is a component module of Spector and requires the main Spector framework for full functionality. Please refer to the main Spector documentation for complete setup and usage instructions.
