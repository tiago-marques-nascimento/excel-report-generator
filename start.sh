#!/bin/bash
cd /home/tiago-nascimento/tiago/workspace/excel-report-generator;
mvn clean install;
cd /home/tiago-nascimento/tiago/workspace/excel-report-generator/target;
java -jar excel-report-generator-1.0.0.jar;
cd /home/tiago-nascimento/tiago/workspace/excel-report-generator;
