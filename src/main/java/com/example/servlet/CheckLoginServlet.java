package com.example.servlet;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import org.json.JSONObject;
import java.io.IOException;

@WebServlet("/checkLogin")
public class CheckLoginServlet extends HttpServlet {
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        HttpSession session = request.getSession(false);
        boolean loggedIn = session != null && session.getAttribute("userId") != null;
        response.setContentType("application/json");
        response.getWriter().write(new JSONObject().put("loggedIn", loggedIn).toString());
    }
}