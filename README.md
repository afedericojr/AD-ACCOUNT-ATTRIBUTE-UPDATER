# PowerShell

```
THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```


This script is for general use of mass updating any Active Directory attribute via a CSV file of desired attributes for each user, so they can all be personalized.

There are vendors that sell solutions like this, but I figured, why not make it. 
This could be an enormous time saver!

I trapped as many errors I could think to test for in the given amount of time.

Notable features:
-	Interactive
-	Updates specified OU
-	Can update any number of users
-	Can update any valid attribute with the desired value
-	Can clear attributes if a NULL entry is desired
-	Confirms and verifies actions
-	Catches errors in:
  -	Missing file or bad file name
  -	Invalid column names
  -	Commands not available
    - Asks if you would like to install and then automatically installs the cmdlet
  -	Invalid attribute name
  -	Any uncaught exception and then provides a description
-	Success and Error log files

Here is a screenshot of it running:

![Screenshot](/AD_Profile_Updater.png?raw=true "Screenshot")
