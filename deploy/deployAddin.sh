
#!/bin/bash

set -e

SCRIPTDIR="$( cd "$(dirname "$0")" >/dev/null 2>&1 ; pwd -P )"

cd $SCRIPTDIR
echo "Deploing to region $REGION"
echo "Installing main modules"
npm install --no-save


function deploy {
    # Build
    start=`date +%s`
    node build.js
    end=`date +%s`
    runtime=$((end-start))
    echo "Addin build done in ${runtime}s"

    # Sync dist to s3
    CLOUD_FRONT_ID=`cat sls-stack-output.json | python3 -c 'import json,sys;obj=json.load(sys.stdin);print(obj["CloudFrontID"])'`
    BUCKET_NAME=`cat sls-stack-output.json | python3 -c 'import json,sys;obj=json.load(sys.stdin);print(obj["SiteBucketName"])'`
    REGION=`cat sls-stack-output.json | python3 -c 'import json,sys;obj=json.load(sys.stdin);print(obj["Region"])'`

    start=`date +%s`
    aws --profile fractal --region $REGION s3 sync --cache-control 'max-age=31536000' --exclude Home.html dist/ s3://$BUCKET_NAME/
    aws --profile fractal --region $REGION s3 sync --cache-control 'no-cache' dist/ s3://$BUCKET_NAME/
    end=`date +%s`
    runtime=$((end-start))
    echo "S3 sync done in ${runtime}s"

    aws cloudfront create-invalidation --profile fractal --distribution-id ${CLOUD_FRONT_ID} --paths '/*'
}

deploy