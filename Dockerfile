FROM python:alpine
LABEL authors="Michael Fretz"

WORKDIR /usr/src/mse
COPY ./requirements.txt ./app.py ./main.py ./google_auth.py ./sheet_info.json ./
COPY ./templates ./templates

RUN apk update && \
    apk upgrade && \
    pip install --upgrade pip && \
    pip install -r ./requirements.txt

CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]