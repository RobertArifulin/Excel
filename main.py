import streamlit as st
from app import App
from pages import comparator, instruction
import pip
pip.main(["install", "openpyxl"])


def main():
    app = App()

    app.add_page("Сравнение Таблиц", comparator)
    app.add_page("Инструкция", instruction)

    app.run()


if __name__ == "__main__":
    main()
