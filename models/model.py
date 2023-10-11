import os


def init_print():
    print("Welcome to use Cx330_502's schedule program, the program running time may be long, please be "
          "patient!")
    print("    ___  _  _  ___  ___   ___       ___   ___  ___  ")
    print("   / __)( \\/ )(__ )(__ ) / _ \\     | __) / _ \\(__ \\ ")
    print("  ( (__  )  (  (_ \\ (_ \\( (_) )___ |__ \\( (_) )/ _/ ")
    print("   \\___)(_/\\_)(___/(___/ \\___/(___)(___/ \\___/(____)")
    print()


def input_model():
    print("Please input the model you want to use:")
    print("1. input and output both in current directory")
    print("2. './data/cpp_dependency' for input and './output/cpp_dependency' for output")
    print("3. Customizing the working directory")
    while True:
        model = input("Please input 1 or 2 or 3: ")
        model = int(model)
        if model == 1:
            input_root0 = "./"
            output_root0 = "./"
            break
        elif model == 2:
            input_root0 = "./data/cpp_dependency"
            output_root0 = "./output/cpp_dependency"
            break
        elif model == 3:
            input_root0 = input("Please input the input directory: ")
            output_root0 = input("Please input the output directory: ")
            break
    os.makedirs(input_root0, exist_ok=True)
    os.makedirs(output_root0, exist_ok=True)
    return input_root0, output_root0
