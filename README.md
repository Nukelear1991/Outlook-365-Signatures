# Outlook-365-Signatures
Two PS scripts. One to create signatures using information from AD, and the other to take those signatures and place them on the client.

Copy the signature creation script to the server you want to host the signature files, be aware that this could be a huge repository of files in a large network.

Create a startup script GPO for all users, and designate the client signature script. This will ensure that the client checks for updated signatures every time the user logs in.
