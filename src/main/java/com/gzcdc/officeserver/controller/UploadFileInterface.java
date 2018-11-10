package com.gzcdc.officeserver.controller;

import com.gzcdc.officeserver.util.OfficeFileUtils;
import org.apache.commons.io.FileUtils;
import org.springframework.stereotype.Controller;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

@Controller
public class UploadFileInterface {

	private final int MAX_SIZE = 1024 * 1024 * 10; // 限制用户头像的最大值为1M
//	private String[] extendNamesArray = { ".jpg", ".jpeg" }; // 用户头像的扩展名数组，方面验证
//	private String rootPath; // 文件根路径
	private String imageNewPath; // 头像新路径（包含头像名以及扩展名）
	private String imageNames; // 头像的新名字（时间+用户名），时间精确到毫秒
	private String extendName; // 头像的扩展名，进行扩展名验证，以达到对用户头像的图片类型限制
	private String message; // 用于返回上传头像的信息
	private String imageURL; // 用于返回用户头像存放的物理路径
	private MultipartFile imageFile;

	@RequestMapping(value = "/uploadHeadPortrait", method = RequestMethod.POST)
	@ResponseBody
	private Map<String, Object> uploadHeadPortrait(HttpServletRequest request) {
		Map<String, Object> map = new HashMap<String, Object>();

		MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;
		// 获取上传头像
		imageFile = multipartRequest.getFile("file");
		// 获取上传头像的文件名
		String fileName = imageFile.getOriginalFilename();
		System.out.println("OriginalFilename:" + fileName);
		// 获取文件扩展名
		extendName = fileName.substring(fileName.lastIndexOf("."));
		// 获取上传头像的大小
		int imageSize = (int) imageFile.getSize();
		// 验证头像的扩展名是否符合要求
		if ((imageSize <= MAX_SIZE)) {
			//rootPath = request.getSession().getServletContext().getRealPath("//upload");
			imageNames = getUploadCurrentTime(); // 重新命名上传头像名称
			imageNewPath = OfficeFileUtils.LINUXROOTPATH + "/upload/" + imageNames + extendName;
			// 判断新路径是否等于数据库中已存在的路径，不等于，则存储新路径，删除原有头像文件
			if (imageSave(imageNewPath)) {
				try {
//					tabPictureService.insertTabPicture(picture);
				} catch (Exception e) {
					e.printStackTrace();
					message = "文件信息保存失败";
					imageURL = null;
					map.put("resultMsg", message);
					map.put("data", imageURL);
					map.put("isSuccess", false);
					return map;
				}
				message = "文件上传成功";
				imageURL = imageNewPath.substring(imageNewPath.indexOf("/upload"));
				map.put("resultMsg", message);
				map.put("data", imageURL);
				map.put("isSuccess", true);
			} else {
				// 保存失败
				message = "文件保存失败";
				imageURL = null;
				map.put("resultMsg", message);
				map.put("data", imageURL);
				map.put("isSuccess", false);
			}

		} else { // 图像格式不符合或者头像的大小大于1M
			message = "文件上传失败，格式或大小不符合";
			imageURL = null;
			map.put("resultMsg", message);
			map.put("data", imageURL);
			map.put("isSuccess", false);
		}
		return map;
	}

	// 获取头上上传的当前时间
	private String getUploadCurrentTime() {
		return new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date());
	}

	// 对原有头像的新路径进行存储，存储后进行检查，时候已经存储，存储成功返回true，失败则返回false
	private boolean imageSave(String imageNewPath) {
		File uploadFile = new File(imageNewPath);
		try {
			if (!uploadFile.exists()) {
				// 先得到文件的上级目录，并创建上级目录，在创建文件
				uploadFile.getParentFile().mkdir();
				// 创建文件
				uploadFile.createNewFile();
			}
			FileCopyUtils.copy(imageFile.getBytes(), uploadFile);
			FileUtils.copyInputStreamToFile(imageFile.getInputStream(), new File(imageNewPath));
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}

	}
}
