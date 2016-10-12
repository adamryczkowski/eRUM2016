Seamless external R server integration with Excel with step-by-step debugging of the R code
===========================================================================================


!

## Demonstration

!

# How does it work?

> It communicates with RServe via Excesi-R using VBA macros. For debug the macros are further forwarding the commands to the
> `svSocket` server. `svSocket` server allows for seamless R session sharing because it runs the separate (`tcltk`) event loop.

!

# Client requirements

* MS Windows (administrative rights are *not required*)
* **32bit** MS Office

!

# Server requirements

* R, accessible without proxy to the Windows and with opened port
* Recent versions of `RServe` and `svSocket`
* For best experience use RStudio (**not RStudio-server**)

!

# Limitations

* If not Linux/Mac then there is no support background task execution (which is implemented via forks)
* Connection between Excel and R is not secure! Use it over trusted network

!

# Limitations (part 2)

* svSocket client (e.g. RStudio) must be running on the same machine as RServe server - this is the limitation of the svSocket library, also because of security reasons.  

!

# Limitations (part 3)

* Because VBA is single-threaded, if you send a long-running R command synchronously, it will block Excel as well. Use asynchronous commands!

!

# Features

* Asynchronous execution
* Nice API for your VBA projects

!

## How to setup

### Excel client

`1`. Download Excelsi-R.dll from Excelsi-R sourceforge or from https://github.com/adamryczkowski/eRUM2016 

> you actually download the whole bundle `Excelsi-R v0.8 setup.exe`. Just extract the dll from there.

!

### Excel client (part 2)

`2`. Put the file in the location, where you have 'execute' rights. E.g. Desktop.

`3`. Download the `template.xlsm` from my github. Or just the VBA code.

!

### Server part (part 1)

`1`. Run R

`2`. Install `Rserve` and `svSocket`:

> (`install.packages(c('Rserve', 'svSocket'))`)

!

### Server part (part 2)

`3`. Run Rserve server:

> (`Rserve::run.Rserve()`)

!

### Server part (part 3)

`4`. Launch RStudio or just a plain R session on the same machine.

`5`. In that session run:

> `svSocket::startSocketServer()`


### Source code

https://github.com/adamryczkowski/eRUM2016

