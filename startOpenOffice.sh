#!/bin/bash
while true; do
	/usr/lib/openoffice/program/soffice -headless "-accept=socket,host=localhost,port=8100;urp" -nologo -nofirststartwizard
	sleep 1
done