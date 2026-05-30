import { Router, type IRouter } from "express";
import healthRouter from "./health";
import pushStatusRouter from "./push-status";

const router: IRouter = Router();

router.use(healthRouter);
router.use(pushStatusRouter);

export default router;
