FROM python:3
RUN pip install xlrd && pip install Pillow
COPY . /app
WORKDIR /app
CMD python gerador_certificados.py