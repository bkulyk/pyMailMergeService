#!/bin/bash
while true; do
        LIBRE="/usr/lib/libreoffice/program/soffice"
        OPENO="/usr/lib/openoffice/program/soffice"
        if [ -e "$OPENO" ]
        then
                OFFICE="$OPENO"
        fi
        if [ -e "$LIBRE" ]
        then
                OFFICE="$LIBRE"
        fi
        $OFFICE -headless "-accept=socket,host=localhost,port=8100;urp" -nologo -nofirststartwizard
        sleep 1
done

