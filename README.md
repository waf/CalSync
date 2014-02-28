CalSync
=======

Cross-domain Exchange calendar synchronization usually requires IT setup/coordination, which isn't always possible. This utility provides an alternative way to achieve calendar synchronization. This tool assumes that you have two computers, each with separate Exchange accounts, and you want to keep your free/busy status synchronized.

Configuration Options
=====================

- TargetEmailAddress - The email address of the exchange account you want to send appointments to
- SyncRangeDays - The number of days (from the current day) to synchronize.
- Send - true/false - Should this computer send local calendar appointments to the target email address?
- Receive - true/false - Should this computer process incoming calendar appointments and add them to the local calendar?

Install / Setup
===============

To set up two-way sync between calendars on two separate computers (computer A and computer B):

1. Download [CalSync.exe](https://github.com/waf/CalSync/raw/master/binaries/CalSync.exe) and the configuration file [CalSync.exe.config](https://github.com/waf/CalSync/raw/master/binaries/CalSync.exe.config) to both Computer A and B.
2. Set your TargetEmailAddress in the config files on each computer. Computer A's config file should use Computer B's email address, and vice versa.
3. Run CalSync.exe on both computers to perform some required set-up. 
4. (Optional) Set up a Windows Scheduled Task to run CalSync.exe periodically.
