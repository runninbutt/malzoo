FROM python:2
WORKDIR /opt/malzoo
COPY . .
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 1338
CMD [ "python", "./malzoo.py" ]