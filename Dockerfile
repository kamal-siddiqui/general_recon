FROM python:3.8

RUN apt-get update && apt-get install --no-install-recommends -y build-essential gcc g++ net-tools telnet && \
    addgroup python && adduser python --ingroup python && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /opt    
ENV TZ=Asia/Kolkata
COPY . /opt/
RUN pip install --no-cache-dir -r requirements.txt && chown -R python:python /opt /tmp 
USER python
EXPOSE 3001

CMD [ "python", "/opt/app.py" ]



#FROM python:3.8.16-bullseye
#WORKDIR /app
#RUN apt-get update && apt-get install --no-install-recommends -y build-essential gcc g++
#COPY requirements.txt .
#RUN pip install --no-cache-dir -r requirements.txt
#COPY . .
#CMD [ "python", "app.py" ]