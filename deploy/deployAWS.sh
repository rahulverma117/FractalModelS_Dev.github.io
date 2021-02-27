#!/bin/bash

REGION=us-east-1

SCRIPTDIR="$( cd "$(dirname "$0")" >/dev/null 2>&1 ; pwd -P )"

cd $SCRIPTDIR
echo "Deploing to region $REGION"
echo "Installing main modules"
npm install --no-save

# AWS
start=`date +%s`
./node_modules/.bin/sls deploy --region $REGION --aws-profile fractal
SLS_ERROR=$?
end=`date +%s`
runtime=$((end-start))
echo "Serverless deploy done in ${runtime}s"

if [ $SLS_ERROR -ne 0 ]; then
  exit $SLS_ERROR
fi
