<tr>
  <td style="padding-bottom:15px;"><a href="http://intranet/"><img src="http://intranet/images/yamaha_logo.jpg" alt="home" border="0" /></a></td>
</tr>
<tr>
  <td height="26"><table width="100%" border="0" cellspacing="0" cellpadding="0" height="55">
      <tr>
        <% if strSection = "home" then %>
        <td width="78"><img src="http://intranet/images/btn_home_selected.jpg" alt="home" width="78" height="55" border="0" /></td>
        <% else %>
        <td width="78"><a href="http://intranet/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_home','','http://intranet/images/btn_home_roll.jpg',1)"><img src="http://intranet/images/btn_home.jpg" alt="home" name="home" width="78" height="55" border="0" id="btn_home" /></a></td>
        <% end if
		   if strSection = "av" then %>
        <td width="116"><img src="http://intranet/images/btn_av_selected.jpg" width="116" height="55" border="0" /></td>
        <% else %>
        <td width="116"><a href="http://intranet/divisions/avit/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_av','','http://intranet/images/btn_av_roll.jpg',1)"><img src="http://intranet/images/btn_av.jpg" alt="audio visual" name="btn_av" width="116" height="55" border="0" id="btn_av" /></a></td>
        <% end if 
		   if strSection = "mpd" then %>
        <td width="113"><img src="http://intranet/images/btn_mpd_selected.jpg" alt="MPD" border="0" /></td>
        <% else %>
        <td width="113"><a href="http://intranet/divisions/mpd/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_pro','','http://intranet/images/btn_mpd_roll.jpg',1)"><img src="http://intranet/images/btn_mpd.jpg" alt="MPD" name="btn_pro" border="0" id="btn_pro" /></a></td>
        <% end if		   
		   if strSection = "education" then %>
        <td width="101"><img src="http://intranet/images/btn_ed_selected.jpg" alt="education centre" width="101" height="55" border="0" /></td>
        <% else %>
        <td width="101"><a href="http://intranet/divisions/ymec/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_ed','','http://intranet/images/btn_ed_roll.jpg',1)"><img src="http://intranet/images/btn_ed.jpg" alt="education centre" name="btn_ed" width="101" height="55" border="0" id="btn_ed" /></a></td>
        <% end if
		   if strSection = "corp" then %>
        <td width="120"><img src="http://intranet/images/btn_corp_selected.jpg" alt="corporate communications" width="120" height="55" border="0" /></td>
        <% else %>
        <td width="120"><a href="http://intranet/divisions/corpcomms/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_corp','','http://intranet/images/btn_corp_roll.jpg',1)"><img src="http://intranet/images/btn_corp.jpg" alt="corporate communications" name="btn_corp" width="120" height="55" border="0" id="btn_corp" /></a></td>
        <% end if
		   if strSection = "hr" then %>
        <td width="60"><img src="http://intranet/images/btn_hr_selected.jpg" align="human resources" width="60" height="55" border="0" /></td>
        <% else %>
        <td width="60"><a href="http://intranet/divisions/hr/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_hr','','http://intranet/images/btn_hr_roll.jpg',1)"><img src="http://intranet/images/btn_hr.jpg" alt="human resources" name="btn_hr" width="60" height="55" border="0" id="btn_hr" /></a></td>
        <% end if 
		   if strSection = "services" then %>
        <td width="91"><img src="http://intranet/images/btn_services_selected.jpg" alt="services" width="91" height="55" border="0" /></td>
        <% else %>
        <td width="91"><a href="http://intranet/divisions/service/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_services','','http://intranet/images/btn_services_roll.jpg',1)"><img src="http://intranet/images/btn_services.jpg" alt="services" name="btn_services" width="91" height="55" border="0" id="btn_services" /></a></td>
        <% end if 
		   if strSection = "logistics" then %>
        <td width="93"><img src="http://intranet/images/btn_logistics_selected.jpg" alt="logistics" width="93" height="55" border="0" /></td>
        <% else %>
        <td width="93"><a href="http://intranet/divisions/logistics/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_logistics','','http://intranet/images/btn_logistics_roll.jpg',1)"><img src="http://intranet/images/btn_logistics.jpg" alt="logistics" name="btn_logistics" width="93" height="55" border="0" id="btn_logistics" /></a></td>
        <% end if 
		   if strSection = "finance" then %>
        <td width="89"><img src="http://intranet/images/btn_finance_selected.jpg" alt="finance" width="89" height="55" border="0" /></td>
        <% else %>
        <td width="89"><a href="http://intranet/divisions/finance/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_finance','','http://intranet/images/btn_finance_roll.jpg',1)"><img src="http://intranet/images/btn_finance.jpg" alt="finance" name="btn_finance" width="89" height="55" border="0" id="btn_finance" /></a></td>
        <% end if 
		   if strSection = "it" then %>
        <td width="54"><img src="http://intranet/images/btn_it_selected.jpg" alt="information technology" width="54" height="55" border="0" /></td>
        <% else %>
        <td width="54"><a href="http://intranet/divisions/it/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_it','','http://intranet/images/btn_it_roll.jpg',1)"><img src="http://intranet/images/btn_it.jpg" alt="information technology" name="btn_it" width="54" height="55" border="0" id="btn_it" /></a></td>
        <% end if 
		   if strSection = "videos" then %>
        <td width="54"><img src="http://intranet/images/btn_videos_selected.jpg" alt="videos" width="84" height="55" border="0" /></td>
        <% else %>
        <td width="54"><a href="http://intranet/videos/" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('btn_videos','','http://intranet/images/btn_videos_roll.jpg',1)"><img src="http://intranet/images/btn_videos.jpg" alt="videos" name="btn_videos" width="84" height="55" border="0" id="btn_videos" /></a></td>
        <% end if %>
        <td class="empty_btn_nav">&nbsp;</td>
      </tr>
    </table></td>
</tr>
