# Enhancing Test Data Generation via Path-Grouped Reusable Prioritized DQN and MI-PSO for Mutation Testing

This repository provides the source code for the study **"Enhancing Test Data Generation via Path-Grouped Reusable Prioritized DQN and MI-PSO for Mutation Testing"**.

The project implements a hybrid test-data generation framework for mutation testing and path coverage. The core method combines a **Path-Grouped Reusable Prioritized Deep Q-Network (PRP-DQN)** with a **Mutation-Inversion Particle Swarm Optimization (MI-PSO)** algorithm. PRP-DQN improves sample reuse and training efficiency across correlated paths, while MI-PSO helps the population escape local optima during path-oriented search.

## Repository Structure

```text
Core Algorithm Implementations/
|-- PRP-DQN/
|   |-- Algorithm 1pathgrouping.py
|   |-- Algorithm 2sample_selection.py
|   |-- Generate test data.py
|   `-- code.py
`-- Experiment code/
    |-- Experiment One.py
    |-- Experiment Two No grouping.py
    |-- Experiment Two groups without model reuse.py
    |-- Experiment Two Grouping and Model Reuse.py
    |-- Experiment Two Random grouping and model reuse.py
    |-- Experiment Three No priority.py
    |-- Experiment Three Priority Granted.py
    |-- Experiment Four DQN.py
    |-- Experiment Four PPO.py
    |-- Experiment Four SAC.py
    |-- Experiment Five PSO.py
    |-- Experiment Five DQN+PSO.py
    |-- Experiment Five PRPDQN+PSO.py
    `-- Experiment Five PRPDQN+MIPSO.py
```

## Environment

The experiments were developed and tested in the following environment:

- Operating system: Windows 11, 64-bit
- Python: 3.8 or later
- CPU: Intel Core i5 or higher
- Memory: 16 GB RAM
- Storage: 512 GB SSD or higher

Install the required Python packages:

```bash
pip install numpy pandas torch openpyxl psutil tqdm
```

If GPU acceleration is required, install the PyTorch build that matches your CUDA environment.

## Core Method

The proposed method contains four main components.

1. **Path grouping:** Target paths are grouped according to Jaccard similarity. Highly correlated paths and low-correlation paths are trained in separate stages.
2. **High-quality sample selection:** Candidate samples are ranked using path similarity, path-length difference, robustness, and, for low-correlation paths, complementary Q-value information.
3. **PRP-DQN training:** Prioritized replay is used to reuse high-value transitions. The model trained on highly correlated paths can be reused and fine-tuned for low-correlation paths.
4. **MI-PSO optimization:** Mutation-inversion PSO improves diversity when population stagnation is detected and supports difficult path coverage.

## Main Scripts

### PRP-DQN implementation

- `PRP-DQN/Algorithm 1pathgrouping.py`: path grouping based on path correlation.
- `PRP-DQN/Algorithm 2sample_selection.py`: high-quality sample selection.
- `PRP-DQN/Generate test data.py`: generation of scored test data.
- `PRP-DQN/code.py`: complete PRP-DQN and MI-PSO workflow.

Run the full PRP-DQN workflow:

```bash
cd "Core Algorithm Implementations/PRP-DQN"
python code.py
```

Run individual modules:

```bash
python "Algorithm 1pathgrouping.py"
python "Algorithm 2sample_selection.py"
python "Generate test data.py"
```

### Comparative experiments

The `Experiment code/` directory contains the scripts used for the ablation studies and baseline comparisons reported in the experiments. Each file is self-contained and can be executed independently.

| Experiment | Script | Purpose |
|---|---|---|
| Experiment One | `Experiment One.py` | Evaluates isolated paths with a four-criteria scoring strategy based on path similarity, path-length difference, robustness, and DQN-derived information. |
| Experiment Two | `Experiment Two No grouping.py` | Runs the training workflow without path grouping. This script is used as the no-grouping baseline. |
| Experiment Two | `Experiment Two groups without model reuse.py` | Uses path grouping but trains grouped models independently, without transferring model parameters between groups. |
| Experiment Two | `Experiment Two Grouping and Model Reuse.py` | Uses path grouping and model reuse. The model trained on the high-correlation group is reused for the low-correlation group. |
| Experiment Two | `Experiment Two Random grouping and model reuse.py` | Replaces correlation-based grouping with random grouping while keeping model reuse, testing whether the grouping strategy itself contributes to performance. |
| Experiment Three | `Experiment Three No priority.py` | Removes prioritized replay and uses ordinary replay sampling. |
| Experiment Three | `Experiment Three Priority Granted.py` | Enables prioritized replay to evaluate the effect of priority-based experience reuse. |
| Experiment Four | `Experiment Four DQN.py` | Uses DQN as the reinforcement-learning baseline. |
| Experiment Four | `Experiment Four PPO.py` | Uses PPO as the reinforcement-learning baseline. |
| Experiment Four | `Experiment Four SAC.py` | Uses SAC as the reinforcement-learning baseline. |
| Experiment Five | `Experiment Five PSO.py` | Uses standard PSO only. |
| Experiment Five | `Experiment Five DQN+PSO.py` | Combines standard DQN with standard PSO. |
| Experiment Five | `Experiment Five PRPDQN+PSO.py` | Combines PRP-DQN with standard PSO. |
| Experiment Five | `Experiment Five PRPDQN+MIPSO.py` | Runs the proposed PRP-DQN and MI-PSO method. |

Most comparative scripts are configured for 20 independent runs by default through `NUM_RUNS = 20`. The number of runs can usually be changed either by modifying `NUM_RUNS` in the script or by passing a command-line argument when the script supports it.

Example commands:

```bash
cd "Core Algorithm Implementations/Experiment code"

python "Experiment Two Grouping and Model Reuse.py"
python "Experiment Three Priority Granted.py"
python "Experiment Four SAC.py"
python "Experiment Five PRPDQN+MIPSO.py"
```

For quick local testing, reduce `NUM_RUNS`, sample counts, or maximum iterations inside the corresponding script.

## Experimental Settings

### State and Action Space

- The experiments use three-dimensional input states.
- Different experiments use different semantic names for the three variables, such as `(dx, dy, dz)`, `(weather, time_period, z)`, or `(light, temp, moisture)`.
- The action space contains dimension-wise positive and negative perturbations.
- In the PRP-DQN and PSO comparison scripts, dynamic step sizes are generated from the variable ranges using ratios such as 70%, 50%, 20%, 10%, and 5%.

### Experimental Groups

The experimental code is organized around five groups:

1. **Experiment One:** isolated-path scoring and screening.
2. **Experiment Two:** effect of path grouping and model reuse.
3. **Experiment Three:** effect of prioritized replay.
4. **Experiment Four:** comparison with DQN, PPO, and SAC reinforcement-learning baselines.
5. **Experiment Five:** comparison among PSO, DQN+PSO, PRPDQN+PSO, and PRPDQN+MIPSO.

This organization matches the file names in `Experiment code/` so that each experiment can be reproduced directly from its corresponding script.

### PRP-DQN Hyperparameters

- Optimizer: Adam
- Learning rate: 0.001
- Discount factor: `gamma = 0.99`
- Initial exploration rate: `epsilon = 1.0`
- Epsilon decay: `0.995`
- Minimum epsilon: `0.1`
- Batch size: `32`
- Prioritized replay exponent: `alpha = 0.6`

### MI-PSO Hyperparameters

- Swarm size: 20
- Maximum iterations: 3000
- Inertia weight: `w = 0.7`
- Acceleration coefficients: `c1 = 1.5`, `c2 = 1.5`
- Maximum velocity: dynamically limited according to the variable range
- Local-optimum detection threshold: coefficient-of-variation threshold `CV_Threshold = 1.2`

## Running the Experiments

### Run the proposed PRP-DQN + MI-PSO method

```bash
cd "Core Algorithm Implementations/Experiment code"
python "Experiment Five PRPDQN+MIPSO.py"
```

### Run PSO-related baselines

```bash
python "Experiment Five PSO.py"
python "Experiment Five DQN+PSO.py"
python "Experiment Five PRPDQN+PSO.py"
```

### Run reinforcement-learning baselines

```bash
python "Experiment Four DQN.py"
python "Experiment Four PPO.py"
python "Experiment Four SAC.py"
```

### Run ablation experiments

```bash
python "Experiment Two No grouping.py"
python "Experiment Two groups without model reuse.py"
python "Experiment Two Grouping and Model Reuse.py"
python "Experiment Two Random grouping and model reuse.py"
python "Experiment Three No priority.py"
python "Experiment Three Priority Granted.py"
```

These scripts print progress information to the command line and may also generate Excel summaries, CSV files, model parameter files, or path-sample files depending on the experiment.

## Outputs

Depending on the script, outputs may include:

- generated path samples,
- trained model parameter files,
- CSV summaries,
- Excel reports generated with `openpyxl`,
- console logs for path coverage, similarity, runtime, and optimization statistics.

Output directories are defined inside the corresponding scripts. Please adjust paths if running the code on a different machine.

## Notes for Reproducibility

- Some experiments use random sampling and stochastic optimization. For strict reproducibility, set the random seeds in the corresponding scripts before running.
- Long-running experiments may require several minutes or more depending on hardware.
- The released scripts are intended to preserve the experimental logic used in the study. Users may adjust output paths, run counts, and sample sizes for local testing.

## Citation

If you use this code in your research, please cite the corresponding thesis or paper:

```text
Enhancing Test Data Generation via Path-Grouped Reusable Prioritized DQN and MI-PSO for Mutation Testing.
```

## License

Please refer to the repository license file. If no license file is provided, contact the authors before redistribution or commercial use.
