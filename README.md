Exchange-Find-Contacts
======================

A small C# application that Sync data from a Exchange Server and push's it to a database (To keep an audit on contacts)

The application will keep a record of all the users contacts in its Exchange Server. It will run every 15minutes, or you could add an Thread to keep the application running.

It will keep checking if there has been an update/delete/add of contacts and then log the change in a database and keeps a track of all the changes in an audit file.

Too use the application you will have to change the App.config to point to your Exchange Server and your database connection. You will also have to run the SQL File found in the SQL folder.  
