# -----------------------------------------------------------
#  Grant Service Control Permissions Script Using PowerShell
# -----------------------------------------------------------
# Version: 1.0
#
# Last Updated 2024-11-01 
#
# PURPOSE: This script is intended to be used during system deployment to 
#          edit the access control list for a Windows Service and allow
#          the specified user or group to control it using the standard
#          Windows Service control metods and APIs.
#          Please read and acknowledge the license below before using.
#
# Acknowledgements: The WebMD Health Services Windows Deployment Team for
#                   creating a powershell module that is easily installed, 
#                   executed and removed from a target system which allows
#                   extensive management and control of system properties.
#                   Please review the documentation at https://get-carbon.org/
#                   and the code at: https://github.com/webmd-health-services/Carbon
#
# --------------------------------------------------------------------------------
# SOFTWARE LICENSE:
#
# MIT License
#
# Copyright (c) 2024 erikpt
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
# --------------------------------------------------------------------------------


# Script Variables Requiring User Input:
[string] $ServiceToModify = "TestService"

# Script Variables Set Automatically
[string] $UserOrGroupToGrant = $ENV:ComputerName + "\Users"


# Download and Install WebMD's Carbon powershell module
# Documentation: https://github.com/webmd-health-services/Carbon
Install-Module Carbon -Scope AllUsers -AcceptLicense -Force -SkipPublisherCheck -Confirm:$false

# Import the Carbon Module so we can use it's commands
Import-Module Carbon -Force

# Set the permissions on the target service so that logged-on users can control the service state including Start, Stop, Restart.
Grant-ServiceControlPermission -ServiceName $ServiceToModify -Identity $UserOrGRoupToGrant -Confirm:$false

# Remove the Carbon System Configration Power Shell Module
Uninstall-Module Carbon -AllVersions -Force -Confirm:$false
