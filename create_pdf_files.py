import pdfkit




pdfkit.from_file(
    "badanie.htm",
    "badanie.pdf",
    verbose=True,
    options={"enable-local-file-access": True},
)