process.env.NODE_ENV = "test";
process.env.JWT_SECRET =
  process.env.JWT_SECRET ||
  "test-secret-64-chars-long-at-least-for-hs256-validation";
process.env.DATABASE_URL =
  process.env.DATABASE_URL ||
  "postgresql://excelagent:dev_password_change_me@localhost:5432/excel_agent";
process.env.YOOKASSA_SHOP_ID = process.env.YOOKASSA_SHOP_ID || "test_shop";
process.env.YOOKASSA_SECRET_KEY =
  process.env.YOOKASSA_SECRET_KEY || "test_secret";
