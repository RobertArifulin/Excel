import streamlit as st
from comparator import start_cmp
from PIL import Image


def comparator():
    st.markdown('# Сравнение таблиц')
    f1, f2 = False, False

    our_register = st.file_uploader("Выбрать внутренний реестр", type=['xlsx', "xls"], key=1)
    if our_register is not None:
        f1 = True

    mavr_register = st.file_uploader("Выбрать реестр от МАВРа", type=['xlsx', "xls"], key=2)
    if mavr_register is not None:
        f2 = True

    if st.button('Сравнить', disabled=not f1 or not f2):
        start_cmp(our_register, mavr_register)


def instruction():
    st.markdown('# Инструкция')
    img1_2 = Image.open('images/1-2.png')
    img3 = Image.open('images/3.png')
    img4 = Image.open('images/4.png')
    img5 = Image.open('images/5.png')
    img6 = Image.open('images/6.png')
    img7 = Image.open('images/7.png')
    img8 = Image.open('images/8.png')
    img9 = Image.open('images/9.png')
    img0 = Image.open('images/10.png')
    img11 = Image.open('images/11.png')
    img12 = Image.open('images/12.png')

    st.image(img1_2, caption='')
    st.markdown("##### Первым выбранным файлом должен быть ваш реестр. Например:")
    st.image(img3, caption='')
    st.markdown("##### Далее нужно создать таблицу excel (.xlsx) для реестра от МАВРа.")
    st.markdown("##### Откройте .docx файл с реестром от МАВРа.")
    st.image(img4, caption='')
    st.markdown("##### В нем скопируйте первую таблицу полностью также, как вы бы копировали текст.")
    st.image(img5, caption='')
    st.markdown("##### Вставьте ее в excel файл.")
    st.image(img6, caption='')
    st.markdown("##### Далее скопируйте вторую таблицу из .docx без первой строчки.")
    st.image(img8, caption='')
    st.markdown("##### Вставьте ее сразу под первой.")
    st.image(img9, caption='')
    st.image(img0, caption='')
    st.markdown("##### Убедитесь, что индексы расставлены правильно.")
    st.image(img7, caption='')
    st.markdown("##### После этого станет доступна кнопка 'Сравнить', нажмите на нее и дождитесь результата.")
    st.image(img11, caption='')
    st.markdown("##### Если все хорошо, то вы увидите кнопку 'Скачать Результат', нажмите на нее и файл с сравнением "
                "загрузится на ваш компьютер.")
    st.image(img12, caption='')

