FROM  python:alpine3.23
WORKDIR /app
COPY . .
RUN pip  install -r  requirements.txt
CMD ["python","app.py"]
