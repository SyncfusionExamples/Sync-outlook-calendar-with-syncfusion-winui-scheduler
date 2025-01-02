# How to Synchronize outlook calendar with syncfusion winui scheduler

This repository contains the samples that demonstrate to Synchronize outlook calendar with Syncfusion WinUI Scheduler

## Syncfusion controls:

This project used the following Syncfusion control(s):
* [SfScheduler](https://www.syncfusion.com/winui-controls/scheduler)

## Requirements to run the sample

Refer to the following link for more details - [System Requirements](https://help.syncfusion.com/winui/system-requirements)

## How to run the sample

1. Clone the sample and open it in Visual Studio.

   *Note: If you download the sample using the "Download ZIP" option, right-click it, select Properties, and then select Unblock.*

2. Register your license key in the App.xaml.cs file as demonstrated in the following code.

		public App()
		{
			//Register Syncfusion license
			Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR LICENSE KEY");

			InitializeComponent();

			MainPage = new App1.MainPage();
		}
		
	Refer to this [link](https://help.syncfusion.com/winui/licensing/overview) for more details.

3. Clean and build the application.

4. Run the application.

## License

Syncfusion has no liability for any damage or consequence that may arise by using or viewing the samples. The samples are for demonstrative purposes, and if you choose to use or access the samples, you agree to not hold Syncfusion liable, in any form, for any damage that is related to use, for accessing, or viewing the samples. By accessing, viewing, or seeing the samples, you acknowledge and agree Syncfusion’s samples will not allow you seek injunctive relief in any form for any claim related to the sample. If you do not agree to this, do not view, access, utilize, or otherwise do anything with Syncfusion’s samples. 
