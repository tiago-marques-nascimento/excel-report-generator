#!/bin/bash
cd /home/tiago-nascimento/tiago/workspace/java-helloworld;
mvn clean install;
cd /home/tiago-nascimento/tiago/workspace/java-helloworld/target;
java -jar helloworld-1.0.0.jar;
cd /home/tiago-nascimento/tiago/workspace/java-helloworld;