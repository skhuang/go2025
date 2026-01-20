from reportlab.pdfgen import canvas

def create_test_pdf(filename):
    c = canvas.Canvas(filename)
    c.drawString(100, 750, "Hello World - Page 1")
    c.showPage()
    c.drawString(100, 750, "Hello World - Page 2")
    c.save()

if __name__ == "__main__":
    create_test_pdf("test.pdf")
    print("Created test.pdf")
