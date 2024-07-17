package com.ek.project.common;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * 登录验证
 *
 * @author ruoyi
 */
@Controller
public class LoginController {



    @GetMapping("/login")
    public String login(HttpServletRequest request, HttpServletResponse response) {

        return "login";
    }


    @GetMapping("/unauth")
    public String unauth() {
        return "error/unauth";
    }


}