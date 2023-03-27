import xlwings as xw
from fastapi import Body, FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pynput.keyboard import Key, Controller

app = FastAPI()


@app.post("/hello")
def hello(data: dict = Body):
    # Instantiate a Book object with the deserialized request body
    book = xw.Book(json=data)

    # Use xlwings as usual
    sheet = book.sheets[0]
    cell = sheet["K1"]
    if cell.value == true:
        keyboard.press(Key.ctrl.value)
        keyboard.press('c')
        keyboard.release('c')
        keyboard.release(Key.ctrl.value)


    # Pass the following back as the response
    return book.json()


# Excel on the web requires CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins="*",
    allow_methods=["POST"],
    allow_headers=["*"],
)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
