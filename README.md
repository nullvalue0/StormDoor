# StormDoor
StormDoor is a BBS door game development kit. It was developed in Visual Basic 6.0 and builds to an OCX control. It was fairly complete including methods to generate ANSI colors, read dropfile properties, respond to user input, and fire different events. 

There are also a number of sample door games included showing usage of various features, including:
- Sending different color foreground/background text
- Reading & writing data from a database
- Menu structures and handling user input
- Paper Rock Scissors - a working multinode interactive game
- TelDoor - a simple door used to telnet to another server 

## History
I started working on this in 2004 to learn more about BBS Door development and create my own utilities and games for the BBSmates BBS. This was developed in Visual Basic 6.0. I recently discovered a backup of this old code. I am releasing it to GitHub mainly so I never lose it again, but anyone is free to do whatever they want with it.

## Control Usage

### Properties

***Alias** (string)*

Holds the client's alias, as read from the dropfile.


***BaudRate** (string)*

Holds the client's baud rate, as read from the dropfile.


***BBSID** (string)*

Holds the client's BBS record ID, as read from the dropfile.


***CommType** (integer)

Holds the client's comm type, as read from the dropfile. Should always be set to 2, because StormDoor only supports telnet.


***Emulation** (integer)*

Holds the client's emulation, as read from the dropfile.


***RealName** (string)*

Holds the client's real name, as read from the dropfile.


***SecurityLevel** (integer)*

Holds the client's security level, as read from the dropfile.


***ThisNode** (integer)*

Holds the client's current BBS node, as read from the dropfile.


***TimeLeft** (integer)*

Holds the client's time left in minutes, as read from the dropfile.


***SocketHandle** (integer)*

Holds the client's winsock connection socket handle, as read from the dropfile.


***UserRecord** (integer)*

Holds the client's BBS record number, as read from the dropfile.


***Echo** (boolean)*

If set to true, StormDoor will automatically echo the client's input. This means the client's terminal should have local_echo turned off.


### Methods

**OpenDropFile(path)**

Opens the door32 dropfile and establishes all connections to the client.


*path* - Path to the door32 dropfile.


**SendNode(tonode, data)**

Sends a message to another node.


*tonode* - Node number of user to send the message to.

*data* - Message to send to the other node.


**Display(data)**

Prints the desired text to the client's terminal.

Use the ChangeColor function to print in desired ANSI colors.


*data* - Text to be displayed on the client's terminal.


**ClearDisplay()**

Sends the ANSI command to clear the screen.


**ClearToEndOfLine()**

Sends the ANSI command to clear from the current position to the end of the line.


**ChangeColor(foreground, background)**

Selects the ANSI colors to be used when the Display function is called.


*foreground* - Foreground color, choose from the popup list.

*background* - Background color, choose from the popup list.


**DisplayANSI(path)**

Displays an ANSI file to the client's terminal.


*path* - Path to the ANSI file.


**Quit()**

This function should be called in order to close the connection gracefully.


**WaitForKey([WaitKeys], [Default]. [TimeLimit])**

Waits for keyboard input from the client. Function returns the key that was pressed.


*WaitKeys* - Optional. Comma-seperated string of valid characters to wait for. ie: "A,B,C". Focus will be return only after one of these keys have been selected. If left blank, focus will be returned after any key is pressed.


*Default* - Optional. Character to be automatically selected if the Enter key is pressed.


*TimeLimit* - Optional. Time in minutes to wait at the prompt. If time limit expires, function returns an empty string.


**MoveToPos(line_num,col_num)**

Sends the ANSI command to move the cursor to the specified coordinates.


*line_num* - Specifies which line number to move to.

*col_num* - Specifies which column number to move to.


### Events

**UserInput(data)**

Event fires when the the client types something from their terminal.

*data* - Contains the string data that was typed at the client terminal.

**ConnectionClosed()**

Event fires when the client's connection is dropped unexpectedly.

**NodeRecv(fromnode,data)**

Event fires when the current node receives a message from another node.


*fromnode* - Contains the node number that sent the message.

*data* - Contains the message that was sent.


**NodeSignOn(node,alias)**

Event fires when another user signs onto a different node in the game.


*node* - Contains the node number that the other user signed onto.

*alias* - Contains the alias of the user that just signed on.


**NodeSignOff(node,alias)**

Event fires when another user signs out of a different node in the game.


*node* - Contains the node number that the other user off of.

*alias* - Contains the alias of the user that just signed off.


**RanOutofTime()**

Event fires when the client runs out of online time. StormDoor internally keeps track of how much time the client has left online, based on the contents of the doorfile.
