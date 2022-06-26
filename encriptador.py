from pyAesCrypt import encryptFile

password = "$$H7fDpjk&prJn3$"
encryptFile("Usuarios.xlsx", "Usuarios.xlsx.aes", password)