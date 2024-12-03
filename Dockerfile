#
# Auteur : sylvain.kieffer@univ-paris13.fr
#
# SPDX-License-Identifier: AGPL-3.0-or-later
# License-Filename: LICENSE
#

FROM python:3.11
COPY . /app
WORKDIR /app
RUN pip install -r requirements.txt
EXPOSE 5000
CMD flask run -h 0.0.0.0 -p 5000