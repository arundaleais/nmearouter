# NMEA Router

Written by Neal Arundale (deceased), this has been uploaded to GitHub so that others can use it.

# Registration

Many thanks to Philippe for decoding the [registration process](https://github.com/arundaleais/nmearouter/blob/master/frmRegister.frm#L334):

* Take your serial number (e.g. 8810)
* Perform the following calculation:
```
serial + (53/serial) + (113*serial/4)
```
  * This is unlikely to be an exact number, you might have to round it (if not sure, try both!)
* Now convert this to hexadecimal
  * In google type in "... to hex" then drop the "0x"
* Finally, add `-3BE` to the end

For example with a serial of `8810`
* `8810 + 53/8810 + 113*8810/4 = 257692.50601`
* We round this up to 257693.
* `257693 to hex` in google = `0x3EE9D`
* So the final registration code is `3EE9D-3BE`
