#!/bin/bash

# Find and replace placeholderss
sed -e "s/\${ADDIN_URL}/$ADDIN_URL/g" -e "s/\${CROSSDOMAIN_URL}/$CROSSDOMAIN_URL/g" /app/manifest.xml.template > /app/manifest.xml


# start the app
npm start