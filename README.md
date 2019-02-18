# SocketTester
A very simple tool to test sockets. Useful when doing socket programming. Capable for both TCP and UDP connections.

# Features

The program uses MDI, which means you are able to run multiple tasks in a single program.

## TCP Client Mode
**Customizable server IP and port**

You can specific the server socket IP and port.

**Socket status display**

You can know the network traffic and the status of the socket. The socket state code can be one of the following:

| Code | Meaning            |
|------|--------------------|
| 0    | Connection Closed  |
| 1    | Socket Open        |
| 2    | Listening          |
| 3    | Connection Pending |
| 4    | Resolving Host     |
| 5    | Host Resolved      |
| 6    | Connecting         |
| 7    | Connected          |
| 8    | Connection Closing |
| 9    | Socket Error       |

**Packet received**

There's a list that shows all packet received. You can select a packet to view its size, text content or binary content. You may save the packet as a file or copy its content to clipboard if you wish.

**Packet sender**

You may make up your packet and send it to the remote socket. Both text mode and binary mode are supported. You may load the content from a file if you wish.

## TCP Server Mode
**Customizable server port**

You can specific the server port to listen on.

**Other features of TCP Server Mode are exactly the same as TCP Client Mode**

## UDP Mode
**Customizable local/remote port and remote IP**

You can specific the local UDP port, the remote UDP socket IP and the remote UDP socket port. Set the Remote IP to your broadcasting IP (e.g. `255.255.255.255`) to broadcast packets.

**Other features of UDP Mode are exactly the same as TCP Client Mode**

# License

MIT
