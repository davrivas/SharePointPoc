using SharepointPoc.Services;
using SharepointPoc.Services.Interfaces;

try
{
    ISharePointService spService = new SharePointService();
    await spService.GetRequest();
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}

