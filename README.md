# .NET COM Test Server

[![Build status](https://ci.appveyor.com/api/projects/status/xoc104acavrbyahu/branch/master?svg=true)](https://ci.appveyor.com/project/jacobsantos/test-com-server-9pdfn/branch/master)

Visual Studio 2013 Community Edition project for creating a COM server for testing in C#.

This should be easier to create using the .NET framework and annotations in C# to test various parts of the COM library.

# Continuous Integration

## Appveyor

The `appveyor.yml` file needs to be edited for it to work on your account.

 1. You must have an [Appveyor account](https://ci.appveyor.com). You can login and create an account using GitHub.
 1. Add the the fork from whatever user or organization you created.
 1. In General Settings:
   1. Check "Ignore appveyor.yml".