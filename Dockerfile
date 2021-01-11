# Create linux image base
FROM ubuntu:20.04

# Define timezone details
ENV TZ=America/Costa_Rica
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime \
	&& echo $TZ > /etc/timezone

# Set up your enviroment
USER root
RUN dpkg --add-architecture i386 && apt-get update \
    && apt-get install -y wine64 \
    && apt-get install -y winetricks 
    #&& chmod +x vcrun6_screen.sh \
    #&& ./vcrun_screen.sh \
    #&& winetricks -q vcrun6 \
    #&& winetricks -q wsh57

# copy vbscript sample files
RUN mkdir -p ~/Desktop/app
COPY . ~/Desktop/app 

# Set up screen permissions 
RUN cd ~/Desktop/app && chmod +x vcrun6_screen.sh \
    && ./vcrun6_screen.sh


	


	


