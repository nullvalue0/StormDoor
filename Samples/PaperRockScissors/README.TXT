This sample shows how to create a simple door game. It stores and 
retreives the scores from a text file, and also allows for 
simultaneous multinode game play.

Just a word of warning: This sample uses the WaitForKey() method.
When this method is called and waiting for the player to press a
key, the server processor usage will go up to 100%. While this
_WILL NOT_ affect server performance, it can look a bit sloppy.
The only way around this is to develop a "state machine" and use
StormDoor's UserInput event, as is shown in the "Database" sample 
code application.

This sample is intended for beginner to intermediate programmers.