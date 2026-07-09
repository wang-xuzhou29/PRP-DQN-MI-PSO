import random
import math

# =================================================================
# Robot path planning and state evaluation simulation
# =================================================================

# --- 1. Constants & Configuration ---
MAX_GRID_SIZE = 500.0
INITIAL_BATTERY = 1000.0
BATTERY_PER_STEP = 1.0
SAFE_DISTANCE = 5.0
CRITICAL_BATTERY_LEVEL = 100.0
MAX_STEP_LENGTH = 40.0
TARGET_X, TARGET_Y, TARGET_Z = 450.0, 450.0, 200.0
MIN_PLANNING_X = 50.0
MIN_PLANNING_Z = 10.0

# --- 2. Robot Status Class ---
class Robot:
    def __init__(self, start_x=0.0, start_y=0.0, start_z=0.0):
        self.x = start_x
        self.y = start_y
        self.z = start_z
        self.battery = INITIAL_BATTERY
        self.status = "Operating"

    def move(self, dx, dy, dz):  # Accept three-dimensional movement as input

        current_step_dist = math.sqrt(dx ** 2 + dy ** 2 + dz ** 2)
        if current_step_dist > MAX_STEP_LENGTH:
            print(f"Warning: Input movement magnitude ({current_step_dist:.2f}m) exceeds the safety threshold!")
            dx, dy, dz = dx * (MAX_STEP_LENGTH / current_step_dist), \
                         dy * (MAX_STEP_LENGTH / current_step_dist), \
                         dz * (MAX_STEP_LENGTH / current_step_dist)

        if self.battery < 0.1:
            self.status = "Dead"
            print(f"[{self.status}] Unable to move, battery power depleted.")
            return

        new_x = self.x + dx
        new_y = self.y + dy
        new_z = self.z + dz

        if abs(dx) < MIN_PLANNING_X:
            print("Path planning warning: X-axis input is too small!")

        if TARGET_Y > self.y and dy < 0.0:
            print("Path planning warning: Y-axis input is opposite to the target direction!")

        if abs(dz) > MIN_PLANNING_Z * 2:
            print("The input movement amount along the Z-axis is extremely large, which may cause instability.")

        distance_moved = math.sqrt((new_x - self.x) ** 2 + (new_y - self.y) ** 2 + (new_z - self.z) ** 2)

        self.x = new_x
        self.y = new_y
        self.z = new_z

        self.x = max(0.0, min(self.x, MAX_GRID_SIZE))
        self.y = max(0.0, min(self.y, MAX_GRID_SIZE))
        self.z = max(0.0, min(self.z, MAX_GRID_SIZE))

        self.battery -= distance_moved * BATTERY_PER_STEP

        if self.battery < CRITICAL_BATTERY_LEVEL:
            self.status = "Low Power"

        if self.battery <= 0.0:
            self.battery = 0.0
            self.status = "Dead"

    def get_distance_to_target(self):
        """Calculate the Euclidean distance to the target (three-dimensional)"""
        return math.sqrt((TARGET_X - self.x) ** 2 + (TARGET_Y - self.y) ** 2 + (TARGET_Z - self.z) ** 2)

    def assess_status(self, obstacle_distance):
        """Based on the assessment of the environment and internal conditions, propose suggestions for the next course of action."""

        if obstacle_distance < SAFE_DISTANCE:
            if obstacle_distance < 1.0:
                self.status = "Stuck"
                return "The obstacle is very close. We need to retreat or find a way around!"
            else:
                return "Obstacle approaching. Please slow down or change direction."

        # if statement 10: Combines internal state and input data (indirectly, via calculated distance & obstacle distance)
        if self.get_distance_to_target() > 200.0 and obstacle_distance > 100.0:
            return "Far from the target and clear environment, full speed ahead!"

        return "Operating normally."

# --- 3. Main Simulation Loop (Generate input data) ---
def main_simulation(steps=10):
    """Main simulation function"""
    robot = Robot()
    print("--- Robot Simulation Started (No-boundary comparison test mode) ---")

    for step in range(1, steps + 1):
        if robot.status in ["Dead", "Reached Target"]:
            print(f"\nSimulation ended at step {step - 1}. Status: {robot.status}")
            break

        print(f"\n--- Step {step} ---")

        # === Dynamic input 1: Obstacle distance (range 1.0 to 500.0) ===
        obstacle_dist = random.uniform(1.0, 500.0)

        # Status evaluation and advice (using input)
        advice = robot.assess_status(obstacle_dist)
        print(f"Obstacle distance: {obstacle_dist:.2f}m. Advice: {advice}")

        # === Dynamic input 2: Movement increments dx, dy, dz (range 1.0 to 50.0) ===
        move_scale = random.uniform(1.0, 50.0)
        dx = move_scale * random.choice([-1, 1])
        dy = move_scale * random.choice([-1, 1])
        dz = move_scale * random.choice([-1, 1])

        # Inject special values to test new if statements (if 3, 4, 5)
        if step == 3:  # Test if 4: Intentionally opposite to target direction (assuming current at 0, target at 450)
            dx, dy, dz = 1.0, -2.0, 1.0  # dy < 0, triggers warning
        if step == 7:  # Test if 3: Intentionally input small dx
            dx, dy, dz = 5.0, 30.0, 5.0  # |dx| < 50.0 (MIN_PLANNING_X)
        if step == 10:  # Test if 5: Intentionally input large dz
            dx, dy, dz = 1.0, 1.0, 30.0  # |dz| > 20.0 (MIN_PLANNING_Z * 2)

        # Execute movement (using input)
        robot.move(dx, dy, dz)

        print(f"Input movement magnitude: ({dx:.2f}, {dy:.2f}, {dz:.2f})")
        print(f"Current position: ({robot.x:.2f}, {robot.y:.2f}, {robot.z:.2f}), Remaining battery: {robot.battery:.2f}%")

    print("\n--- Robot Simulation Ended ---")


# --- 4. Run main function ---
if __name__ == "__main__":
    main_simulation(steps=20)

# =================================================================
# Confirmation of 10 numeric if statements (no-boundary comparison):
# 1. if current_step_dist > MAX_STEP_LENGTH     <-- Dynamic input (dx, dy, dz) compared to performance constant
# 2. if self.battery < 0.1                      <-- Internal state compared to minimum value
# 3. if abs(dx) < MIN_PLANNING_X                <-- Dynamic input (dx) compared to planning constant
# 4. if TARGET_Y > self.y and dy < 0.0          <-- Dynamic input (dy) compared to 0.0 for direction
# 5. if abs(dz) > MIN_PLANNING_Z * 2            <-- Dynamic input (dz) compared to planning constant
# 6. if self.battery < CRITICAL_BATTERY_LEVEL   <-- Internal state compared to critical constant
# 7. if self.battery <= 0.0                     <-- Internal state compared to 0.0
# 8. if obstacle_distance < SAFE_DISTANCE       <-- Dynamic input (obstacle_distance) compared to safety constant
# 9. if obstacle_distance < 1.0                 <-- Dynamic input (obstacle_distance) compared to critical constant
# 10. if self.get_distance_to_target() > 200.0 and obstacle_distance > 100.0
#                                               <-- Internal state compared to constant AND Dynamic input (obstacle_distance) compared to constant
# =================================================================