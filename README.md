usage: cp-multitool.py [-h] [-H API_SERVER] [-U API_USER] [-K] [-P API_PWD] [-C API_CONTEXT] [-p POLNAME] [-f CFILTER] [-o OUTFILE] [-e EXPORTSELECT] [-v] {show,export,test} ...

positional arguments:

  {show,export,test}    Tell what to do (test || show || export)

    show                Print given information / objects and their properties (use with caution)

    export              Save output of given information / objects and their properties to a file

    test                Basic connectivity test and (if successful) fetch some information from server

options:

  -h, --help            show this help message and exit

  -H API_SERVER, --api_server API_SERVER

                        Target Host (CP Management Server)

  -U API_USER, --api_user API_USER

                        API User

  -K, --key             invoke, that password value is an API Key

  -P API_PWD, --api_pwd API_PWD

                        API Users Password, if using API Key, use this without a user

  -C API_CONTEXT, --api_context API_CONTEXT

                        If SmartCloud-1 is used, enter context information here (i.e. bhkjnkm-knjhbas-d32424b/web_api) - for On Prem enter "-C none"

  -p POLNAME, --polname POLNAME

                        when using "show-access-rulebase" a policy name must be given, using this argument

  -f CFILTER, --filter CFILTER

                        filter by string (if applicable)

  -o OUTFILE, --outfile OUTFILE

                        filename, to save the output in

  -e EXPORTSELECT, --exportselect EXPORTSELECT

                        choose which fields to export to csv (define as string with quotes: "name, ip_address, uid, groups")

  -v, --verbose         Run Script with informational (-v) or debugging (-vv) logging output. For troubleshooting purposes.
