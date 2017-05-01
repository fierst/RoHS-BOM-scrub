# RoHS BOM Scrub Tool
---

We frequently need to scrub a bill of materials for RoHS compliance information. Standard procedure was to manually look it up for each part, but the information is disparate and sometimes can take a significant chunk of time for a single component, much less a board with >50 components.


This tool allows the user to load a text file with multiple part numbers and will find and download part information using the [Octopart API](https://octopart.com/api/home)

__Usage:__
`python rohsscrub.py` 
OR 
`python rohsscrub.py [API KEY]`

Once the script starts, it will prompt the user to load a .txt file with component part numbers separated by whitespace (newlines).

Use of the script requires an API key from Octopart, input either at the command line or once running the script. 

Right now error checking is minimal and there are some UI improvements (w/r/t argument parsing, etc.,) that need to be implemented, but for the most part, it serves its purpose...
___
