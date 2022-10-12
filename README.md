# Convert data to PDF Tables using .NET and  the Foxit API

You need the Foxit SDK for .NET. If you don't have it, download it [here](https://developers.foxit.com/download/). You also need to have the .NET SDK 6.0 or later installed on your PC. If you don't, get it [here](https://dotnet.microsoft.com/en-us/download/dotnet/6.0). Dataset is found in the `/data` directory.

## Getting Started

1. Clone the repository and change directory to it: 

   ```shell
   [SSH] git clone git@github.com:IamGideonIdoko/foxit_pdf_table.git
   [HTTPS] git clone https://github.com/IamGideonIdoko/foxit_pdf_table.git
   
   cd foxit_pdf_table
   ```

2. Create `.env` file from the content of `.env.example` and update it with your Foxit API key and serial number.

   ```shell
   cp .env.example .env
   ```

   Get the key and serial number from the `/<sdk-root>/lib/gsdk_key.txt` and `/<sdk-root>/lib/gsdk_sn.txt` files  respectively

3. Copy the `foxit.dll` and `foxit_dotnetcore.dll` library files from either the  **x64_vc15** or **x86_vc15** folder (depending on the architecture you’re targeting) into the “**lib**” directory in the project.

4. Copy the `foxit.dll` library file into the root of the project.

   ```shell
   cp lib/foxit.dll foxit.dll
   ```

5. Install dependencies

   ```shell
   dotnet install
   ```

6. Run

   ```shell
   dotnet run
   ```

   



