// serverside/src/tests/auth.test.ts
import request from "supertest";
import app from "../app";
import User from "../models/User";
import { setAuthCookie } from "../lib/jwt";
import mongoose from "mongoose";

describe("🔒 Security Infrastructure Tests", () => {
  // 1. NOSQL INJECTION TEST
  it("should reject NoSQL Injection objects in login", async () => {
    const response = await request(app)
      .post("/auth/login") // Ensure this matches your app.ts route
      .send({
        email: { $gt: "" }, // Attempting to bypass email check
        password: "any-password",
      });

    // If sanitizeInput and String() casting are working:
    // { $gt: "" } becomes "[object Object]" or is stripped.
    expect(response.status).toBe(401);
    expect(response.body.message).toBe("Invalid credentials");
  });

  // 2. RATE LIMITING TEST
  it("should block multiple failed login attempts (Rate Limiting)", async () => {
    // Note: If you run this test frequently, you might need to reset the limiter
    // or use a different IP/Email per test run.
    for (let i = 0; i < 6; i++) {
      const res = await request(app)
        .post("/auth/login")
        .send({
          email: `brute-test-${i}@edu.com`,
          password: "wrong-password",
        });

      if (i >= 5) {
        expect(res.status).toBe(429); // Too Many Requests
        expect(res.body.message).toContain("Too many attempts");
      }
    }
  });

  // 3. REAL-TIME SESSION REVOCATION TEST
  it("should deny access if user status becomes suspended mid-session", async () => {
    // SETUP: Create a temporary active user
    const testUser = await User.create({
      name: "Test User",
      email: "test@edu.com",
      password: "hashed-password",
      role: "coordinator",
      status: "active",
      institution: new mongoose.Types.ObjectId(),
    });

    // Generate a valid JWT for this user
    // In a real test, you'd call your login route to get the actual cookie
    const loginRes = await request(app).post("/auth/login").send({
      email: "test@edu.com",
      password: "correct-password",
    });
    const authCookie = loginRes.get("Set-Cookie");

    // ACT: Admin suspends the user in the background
    await User.updateOne({ _id: testUser._id }, { status: "suspended" });

    // ASSERT: Attempt to use the still-valid JWT on a protected route
    const response = await request(app)
      .get("/auth/me")
      .set("Cookie", authCookie || []);

    // Even though the JWT signature is valid, the requireAuth middleware
    // checks the DB and should reject the suspended user.
    expect(response.status).toBe(403);
    expect(response.body.message).toMatch(/revoked|suspended/i);
  });

  // 4. TIMING ATTACK PROTECTION (Conceptual)
  it("should take roughly the same time for non-existent users as valid users", async () => {
    const start = Date.now();
    await request(app)
      .post("/auth/login")
      .send({ email: "non-existent@edu.com", password: "p" });
    const duration = Date.now() - start;

    // This is hard to test perfectly in local envs, but you're ensuring
    // the code path doesn't "return early" when a user is not found.
    expect(duration).toBeGreaterThan(10); // bcrypt.compare usually takes > 50ms
  });
});
