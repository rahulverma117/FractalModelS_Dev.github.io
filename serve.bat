@echo off

:check_mkcert:
    if NOT EXIST "mkcert.exe" (
        echo "MkCert binary missing. Donwloading one."
        call :get_mkcert
    ) else call :check_certs
    EXIT /B 0

:get_mkcert
    CScript get_mkcert.wsf
    echo "Got MkCert binary."
    call :check_certs
    EXIT /B 0

:check_certs
    echo "Checking localhost certs."

    if EXIST "localhost_cert\localhost_cert.pem" (
        certutil -f -urlfetch -verify "localhost_cert\localhost_cert.pem"
        if %errorlevel% NEQ 0 (
            echo "Cert invalid. Generating new one."
            call :generate_cert
        ) else (
            echo "Found valid cert."
            call :serve
        )
    ) else (
        echo "Certs missing. Generating one."
        call :generate_cert
    )
    EXIT /B 0

:generate_cert
    mkdir "localhost_cert"

    mkcert -install
    mkcert -key-file "localhost_cert\localhost_key.pem" -cert-file "localhost_cert\localhost_cert.pem" localhost 127.0.0.1
    echo "Generated certs."
    call :serve
    EXIT /B 0

:serve
    python3 serve.py

    EXIT /B 0

call :check_mkcert